
package com.rahul.dailytask.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.rahul.dailytask.entity.WhatsAppBotMonitor;
import com.rahul.dailytask.repository.WhatsAppBotMonitorRepository;

import jakarta.annotation.PostConstruct;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.env.Environment;
import org.springframework.data.domain.Sort;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.List;

@Service
public class WhatsAppBotMonitorService {

    private final WhatsAppBotMonitorRepository repository;
    private final ObjectMapper objectMapper;
    private final Environment environment;
    private final HttpClient httpClient;
    private final TokenRefreshService tokenRefreshService;

    @Value("${whatsapp.monitor.inactive-minutes:15}")
    private long inactiveMinutesLimit;

    private static final ZoneId INDIA_ZONE =
            ZoneId.of("Asia/Kolkata");

    public WhatsAppBotMonitorService(
            WhatsAppBotMonitorRepository repository,
            ObjectMapper objectMapper,
            Environment environment,
            TokenRefreshService tokenRefreshService) {

        this.repository = repository;
        this.objectMapper = objectMapper;
        this.environment = environment;
        this.tokenRefreshService = tokenRefreshService;

        this.httpClient = HttpClient.newBuilder()
                .connectTimeout(Duration.ofSeconds(15))
                .version(HttpClient.Version.HTTP_1_1)
                .build();
    }

    // Load configured clients without overwriting existing database tokens.
    @PostConstruct
    public void loadConfiguredClients() {

        int index = 0;
        int loadedCount = 0;

        // Continue through gaps in configuration indexes.
        while (index < 100) {

            String prefix = "monitor.clients[" + index + "]";

            String name = environment.getProperty(prefix + ".name");
            String apiUrl = environment.getProperty(prefix + ".api-url");
            String configuredToken =
                    environment.getProperty(prefix + ".token");

            index++;

            if (name == null || name.isBlank()
                    || apiUrl == null || apiUrl.isBlank()) {
                continue;
            }

            try {
                WhatsAppBotMonitor client = repository
                        .findByClientNameIgnoreCase(name)
                        .orElseGet(WhatsAppBotMonitor::new);

                client.setClientName(name);
                client.setApiUrl(apiUrl);

                // Use the configured token only when no database token exists.
                if ((client.getToken() == null
                        || client.getToken().isBlank())
                        && configuredToken != null
                        && !configuredToken.isBlank()) {

                    client.setToken(configuredToken.trim());
                }

                client.setActive(true);

                if (client.getStatus() == null) {
                    client.setStatus(WhatsAppBotMonitor.CHECKING);
                }

                repository.save(client);
                loadedCount++;

                System.out.println(
                        "Monitor client loaded: " + name);

            } catch (Exception e) {
                System.err.println(
                        "Could not load client " + name
                                + ": " + e.getClass().getSimpleName());
            }
        }

        System.out.println(
                "Configured monitor clients loaded: " + loadedCount);
    }

    // Check All button: check every active client immediately.
    // Refresh the token if authentication is expired.
    public List<WhatsAppBotMonitor> checkAllClients() {

        List<WhatsAppBotMonitor> clients =
                repository.findByActiveTrue();

        List<WhatsAppBotMonitor> results = new ArrayList<>();

        for (WhatsAppBotMonitor client : clients) {
            try {
                results.add(checkAndRefreshIfExpired(client));

            } catch (Exception e) {
                System.err.println(
                        "Check failed for " + client.getClientName()
                                + ": " + e.getClass().getSimpleName());

                client.setStatus(WhatsAppBotMonitor.API_ERROR);
                client.setLastChecked(
                        LocalDateTime.now(INDIA_ZONE));
                client.setMessage("Monitoring check failed");

                results.add(repository.save(client));
            }
        }

        return results;
    }

    // Check one client's All Chat API and update its status.
    public WhatsAppBotMonitor checkClient(
            WhatsAppBotMonitor client) {

        long startTime = System.currentTimeMillis();
        LocalDateTime checkTime =
                LocalDateTime.now(INDIA_ZONE);

        client.setLastChecked(checkTime);
        client.setStatus(WhatsAppBotMonitor.CHECKING);
        repository.save(client);

        try {
            if (client.getApiUrl() == null
                    || client.getApiUrl().isBlank()) {

                return updateError(
                        client,
                        WhatsAppBotMonitor.API_ERROR,
                        "API URL is missing",
                        null,
                        startTime);
            }

            HttpRequest.Builder builder = HttpRequest.newBuilder()
                    .uri(URI.create(client.getApiUrl()))
                    .timeout(Duration.ofSeconds(30))
                    .header("Accept", "application/json")
                    .header("User-Agent", "DailyTaskManager");

            if (client.getToken() != null
                    && !client.getToken().isBlank()) {

                String token = client.getToken().trim();

                // Avoid adding "Bearer" twice.
                if (token.regionMatches(
                        true, 0, "Bearer ", 0, 7)) {
                    token = token.substring(7).trim();
                }

                builder.header("Authorization", "Bearer " + token);
            }

            HttpRequest request = builder.GET().build();

            HttpResponse<String> response = httpClient.send(
                    request,
                    HttpResponse.BodyHandlers.ofString());

            long responseTime =
                    System.currentTimeMillis() - startTime;

            int statusCode = response.statusCode();

            client.setStatusCode(statusCode);
            client.setResponseTime(responseTime);

            // Only these responses trigger token refresh.
            if (statusCode == 401 || statusCode == 403) {
                return updateError(
                        client,
                        WhatsAppBotMonitor.AUTH_EXPIRED,
                        "Authentication expired or unauthorized",
                        responseTime,
                        startTime);
            }

            if (statusCode < 200 || statusCode >= 300) {
                return updateError(
                        client,
                        WhatsAppBotMonitor.API_ERROR,
                        "API returned HTTP " + statusCode,
                        responseTime,
                        startTime);
            }

            String responseBody = response.body();

            if (responseBody == null || responseBody.isBlank()) {
                return updateError(
                        client,
                        WhatsAppBotMonitor.API_ERROR,
                        "API returned empty response",
                        responseTime,
                        startTime);
            }

            JsonNode root = objectMapper.readTree(responseBody);
            JsonNode resultArray = root.path("data").path("result");

            if (!resultArray.isArray() || resultArray.isEmpty()) {

                client.setStatus(WhatsAppBotMonitor.NO_DATA);
                client.setMessage("No chat activity data available");
                client.setInactiveMinutes(null);
                client.setLastChecked(checkTime);

                return repository.save(client);
            }

            Instant latestActivity = findLatestUpdatedAt(resultArray);

            if (latestActivity == null) {

                client.setStatus(WhatsAppBotMonitor.NO_DATA);
                client.setMessage("No valid updatedAt found");
                client.setInactiveMinutes(null);
                client.setLastChecked(checkTime);

                return repository.save(client);
            }

            long inactiveMinutes = Duration.between(
                    latestActivity,
                    Instant.now()).toMinutes();

            if (inactiveMinutes < 0) {
                inactiveMinutes = 0;
            }

            client.setLastActivity(
                    latestActivity.atZone(INDIA_ZONE)
                            .toLocalDateTime());

            client.setInactiveMinutes(inactiveMinutes);

            if (inactiveMinutes < inactiveMinutesLimit) {
                client.setStatus(WhatsAppBotMonitor.ACTIVE);
                client.setMessage("Bot activity detected");
            } else {
                client.setStatus(WhatsAppBotMonitor.NO_ACTIVITY);
                client.setMessage(
                        "No activity for " + inactiveMinutes + " minutes");
            }

            client.setLastChecked(checkTime);

            return repository.save(client);

        } catch (DateTimeParseException e) {

            return updateError(
                    client,
                    WhatsAppBotMonitor.API_ERROR,
                    "Invalid updatedAt format",
                    null,
                    startTime);

        } catch (InterruptedException e) {

            Thread.currentThread().interrupt();

            return updateError(
                    client,
                    WhatsAppBotMonitor.API_ERROR,
                    "API request interrupted",
                    null,
                    startTime);

        } catch (Exception e) {

            System.err.println(
                    "API check failed for " + client.getClientName()
                            + ": " + e.getClass().getSimpleName());

            return updateError(
                    client,
                    WhatsAppBotMonitor.API_ERROR,
                    "Unable to connect to monitoring API",
                    null,
                    startTime);
        }
    }

    // Shared logic for the manual Check All action and scheduled checks.
    private WhatsAppBotMonitor checkAndRefreshIfExpired(
            WhatsAppBotMonitor client) {

        // Always test the currently stored token against the live API first.
        WhatsAppBotMonitor checked = checkClient(client);

        // Do not refresh for ACTIVE, NO_DATA, NO_ACTIVITY or API_ERROR.
        if (!WhatsAppBotMonitor.AUTH_EXPIRED.equals(
                checked.getStatus())) {
            return checked;
        }

        System.out.println(
                "Authentication expired for "
                        + checked.getClientName()
                        + ". Attempting token refresh.");

        try {
            String newToken = tokenRefreshService.fetchNewToken(checked);

            if (newToken == null || newToken.isBlank()) {

                checked.setStatus(WhatsAppBotMonitor.AUTH_EXPIRED);
                checked.setMessage(
                        "Token refresh failed; previous token retained");

                return repository.save(checked);
            }

            // Store the refreshed token for this client only.
            checked.setToken(newToken.trim());
            checked.setMessage(
                    "Token refreshed; verifying All Chat API");

            repository.save(checked);

            // Verify the new token against the same client's All Chat API.
            WhatsAppBotMonitor verification = checkClient(checked);

            System.out.println(
                    "Token verification for "
                            + checked.getClientName()
                            + ": " + verification.getStatus());

            return verification;

        } catch (Exception e) {

            System.err.println(
                    "Token refresh failed for "
                            + checked.getClientName()
                            + ": " + e.getClass().getSimpleName());

            checked.setStatus(WhatsAppBotMonitor.AUTH_EXPIRED);
            checked.setMessage(
                    "Token refresh or verification failed");

            return repository.save(checked);
        }
    }

    // Run at 12 AM, 6 AM, 9 AM, 10 AM, 7 PM and 10 PM IST.
    // At each scheduled time, check authentication first and refresh
    // only when the All Chat API returns HTTP 401 or 403.
    @Scheduled(
            cron = "0 0 0,6,9,10,19,22 * * *",
            zone = "Asia/Kolkata")
    public void refreshExpiredTokens() {

        boolean enabled = environment.getProperty(
                "whatsapp.monitor.token-refresh-enabled",
                Boolean.class,
                true);

        if (!enabled) {
            System.out.println("Scheduled token checks are disabled.");
            return;
        }

        System.out.println(
                "Scheduled monitoring job started: "
                        + LocalDateTime.now(INDIA_ZONE));

        List<WhatsAppBotMonitor> clients =
                repository.findByActiveTrue();

        for (WhatsAppBotMonitor client : clients) {
            try {
                WhatsAppBotMonitor result =
                        checkAndRefreshIfExpired(client);

                System.out.println(
                        "Scheduled check completed for "
                                + client.getClientName()
                                + ": " + result.getStatus());

            } catch (Exception e) {

                System.err.println(
                        "Scheduled check failed for "
                                + client.getClientName()
                                + ": " + e.getClass().getSimpleName());
            }
        }

        System.out.println("Scheduled monitoring job completed.");
    }

    private Instant findLatestUpdatedAt(JsonNode resultArray) {

        Instant latest = null;

        for (JsonNode item : resultArray) {

            String updatedAt = item.path("updatedAt").asText(null);

            if (updatedAt == null || updatedAt.isBlank()) {
                continue;
            }

            try {
                Instant current = Instant.parse(updatedAt);

                if (latest == null || current.isAfter(latest)) {
                    latest = current;
                }

            } catch (DateTimeParseException e) {
                System.err.println(
                        "Invalid updatedAt value skipped.");
            }
        }

        return latest;
    }

    private WhatsAppBotMonitor updateError(
            WhatsAppBotMonitor client,
            String status,
            String message,
            Long responseTime,
            long startTime) {

        client.setStatus(status);
        client.setMessage(message);
        client.setLastChecked(LocalDateTime.now(INDIA_ZONE));

        if (responseTime != null) {
            client.setResponseTime(responseTime);
        } else {
            client.setResponseTime(
                    System.currentTimeMillis() - startTime);
        }

        return repository.save(client);
    }

    public List<WhatsAppBotMonitor> getAllClients() {
        return repository.findAll(
                Sort.by(Sort.Direction.ASC, "clientName"));
    }

    public List<WhatsAppBotMonitor> getActiveClients() {
        return repository.findByActiveTrue();
    }

    public WhatsAppBotMonitor getClient(Long id) {
        return repository.findById(id)
                .orElseThrow(() ->
                        new RuntimeException(
                                "WhatsApp client not found: " + id));
    }

    public void deleteClient(Long id) {
        WhatsAppBotMonitor client = repository.findById(id)
                .orElseThrow(() ->
                        new RuntimeException(
                                "WhatsApp client not found: " + id));

        repository.delete(client);
    }
}
