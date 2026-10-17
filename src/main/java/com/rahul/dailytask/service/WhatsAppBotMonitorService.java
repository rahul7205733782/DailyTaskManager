
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

    // Load configured clients without overwriting existing DB tokens.
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

                System.out.println("Monitor client loaded: " + name);

            } catch (Exception e) {
                System.err.println(
                        "Could not load client " + name
                                + ": " + e.getClass().getSimpleName());
            }
        }

        System.out.println(
                "Configured monitor clients loaded: " + loadedCount);
    }

    // Check all active clients.
    public List<WhatsAppBotMonitor> checkAllClients() {

        List<WhatsAppBotMonitor> clients =
                repository.findByActiveTrue();

        List<WhatsAppBotMonitor> results = new ArrayList<>();

        for (WhatsAppBotMonitor client : clients) {
            try {
                results.add(checkClient(client));
            } catch (Exception e) {
                System.err.println(
                        "Check failed for " + client.getClientName());

                client.setStatus(WhatsAppBotMonitor.API_ERROR);
                client.setLastChecked(LocalDateTime.now(INDIA_ZONE));
                client.setMessage("Monitoring error");

                results.add(repository.save(client));
            }
        }

        return results;
    }

    // Check one client's All Chat API.
    public WhatsAppBotMonitor checkClient(
            WhatsAppBotMonitor client) {

        long startTime = System.currentTimeMillis();
        LocalDateTime checkTime = LocalDateTime.now(INDIA_ZONE);

        client.setLastChecked(checkTime);
        client.setStatus(WhatsAppBotMonitor.CHECKING);
        repository.save(client);

        try {
            if (client.getApiUrl() == null
                    || client.getApiUrl().isBlank()) {
                return updateError(
                        client, WhatsAppBotMonitor.API_ERROR,
                        "API URL is missing", null, startTime);
            }

            HttpRequest.Builder builder = HttpRequest.newBuilder()
                    .uri(URI.create(client.getApiUrl()))
                    .timeout(Duration.ofSeconds(30))
                    .header("Accept", "application/json")
                    .header("User-Agent", "DailyTaskManager");

            if (client.getToken() != null
                    && !client.getToken().isBlank()) {

                String token = client.getToken().trim();

                if (token.regionMatches(true, 0, "Bearer ", 0, 7)) {
                    token = token.substring(7).trim();
                }

                // Keep this if the API requires Bearer authentication.
                builder.header("Authorization", "Bearer " + token);
            }

            HttpRequest request = builder.GET().build();

            HttpResponse<String> response = httpClient.send(
                    request, HttpResponse.BodyHandlers.ofString());

            long responseTime = System.currentTimeMillis() - startTime;
            int statusCode = response.statusCode();

            client.setStatusCode(statusCode);
            client.setResponseTime(responseTime);

            if (statusCode == 401 || statusCode == 403) {
                return updateError(
                        client, WhatsAppBotMonitor.AUTH_EXPIRED,
                        "Authentication expired or unauthorized",
                        responseTime, startTime);
            }

            if (statusCode < 200 || statusCode >= 300) {
                return updateError(
                        client, WhatsAppBotMonitor.API_ERROR,
                        "API returned HTTP " + statusCode,
                        responseTime, startTime);
            }

            String responseBody = response.body();

            if (responseBody == null || responseBody.isBlank()) {
                return updateError(
                        client, WhatsAppBotMonitor.API_ERROR,
                        "API returned empty response",
                        responseTime, startTime);
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
                    latestActivity, Instant.now()).toMinutes();

            if (inactiveMinutes < 0) {
                inactiveMinutes = 0;
            }

            client.setLastActivity(
                    latestActivity.atZone(INDIA_ZONE).toLocalDateTime());

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
                    client, WhatsAppBotMonitor.API_ERROR,
                    "Invalid updatedAt format", null, startTime);

        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();

            return updateError(
                    client, WhatsAppBotMonitor.API_ERROR,
                    "API request interrupted", null, startTime);

        } catch (Exception e) {
            System.err.println(
                    "API check failed for " + client.getClientName()
                            + ": " + e.getClass().getSimpleName());

            return updateError(
                    client, WhatsAppBotMonitor.API_ERROR,
                    "Unable to connect to monitoring API",
                    null, startTime);
        }
    }

    // Refresh only clients currently marked AUTH_EXPIRED.
    @Scheduled(
            fixedDelayString =
                    "${whatsapp.monitor.token-refresh-fixed-delay:60000}")
    public void refreshExpiredTokens() {

        boolean enabled = environment.getProperty(
                "whatsapp.monitor.token-refresh-enabled",
                Boolean.class,
                false);

        if (!enabled) {
            return;
        }

        System.out.println("Token refresh job started: "
                + LocalDateTime.now(INDIA_ZONE));

        List<WhatsAppBotMonitor> clients =
                repository.findByActiveTrue();

        for (WhatsAppBotMonitor client : clients) {

            if (!WhatsAppBotMonitor.AUTH_EXPIRED.equals(
                    client.getStatus())) {
                continue;
            }

            System.out.println(
                    "Refreshing expired token for "
                            + client.getClientName());

            try {
                String newToken =
                        tokenRefreshService.fetchNewToken(client);

                if (newToken == null || newToken.isBlank()) {
                    client.setStatus(WhatsAppBotMonitor.AUTH_EXPIRED);
                    client.setMessage(
                            "Login failed; previous token retained");
                    repository.save(client);
                    continue;
                }

                // Save only this client's new token.
                client.setToken(newToken.trim());
                client.setMessage(
                        "New token received; verifying All Chat");
                repository.save(client);

                // Verify against this client's own All Chat API.
                WhatsAppBotMonitor result = checkClient(client);

                System.out.println(
                        "Verification result for "
                                + client.getClientName()
                                + ": " + result.getStatus());

            } catch (Exception e) {
                System.err.println(
                        "Refresh failed for " + client.getClientName()
                                + ": " + e.getClass().getSimpleName());

                client.setStatus(WhatsAppBotMonitor.AUTH_EXPIRED);
                client.setMessage(
                        "Token refresh or verification failed");
                repository.save(client);
            }
        }

        System.out.println("Token refresh job completed.");
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
                System.err.println("Invalid updatedAt value skipped.");
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
