package com.rahul.dailytask.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.rahul.dailytask.entity.WhatsAppBotMonitor;
import com.rahul.dailytask.repository.WhatsAppBotMonitorRepository;

import jakarta.annotation.PostConstruct;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.env.Environment;
import org.springframework.data.domain.Sort;
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


    @Value("${whatsapp.monitor.inactive-minutes:15}")
    private long inactiveMinutesLimit;


    private static final ZoneId INDIA_ZONE =
            ZoneId.of("Asia/Kolkata");


    // =========================================================
    // CONSTRUCTOR
    // =========================================================

    public WhatsAppBotMonitorService(
            WhatsAppBotMonitorRepository repository,
            ObjectMapper objectMapper,
            Environment environment) {

        this.repository = repository;

        this.objectMapper = objectMapper;

        this.environment = environment;

        this.httpClient = HttpClient.newBuilder()
                .connectTimeout(Duration.ofSeconds(15))
                .version(HttpClient.Version.HTTP_1_1)
                .build();
    }


    // =========================================================
    // LOAD CLIENTS FROM application.properties
    // =========================================================

    @PostConstruct
    public void loadConfiguredClients() {

        int index = 0;

        while (true) {

            String name =
                    environment.getProperty(
                            "monitor.clients[" + index + "].name"
                    );

            String apiUrl =
                    environment.getProperty(
                            "monitor.clients[" + index + "].api-url"
                    );

            String token =
                    environment.getProperty(
                            "monitor.clients[" + index + "].token"
                    );


            // Stop when next client is not configured

            if (name == null || apiUrl == null) {

                break;
            }


            try {

                WhatsAppBotMonitor client =
                        repository
                                .findByClientNameIgnoreCase(name)
                                .orElseGet(
                                        WhatsAppBotMonitor::new
                                );


                client.setClientName(name);

                client.setApiUrl(apiUrl);

                client.setToken(token);

                client.setActive(true);


                if (client.getStatus() == null) {

                    client.setStatus(
                            WhatsAppBotMonitor.CHECKING
                    );
                }


                repository.save(client);


                System.out.println(
                        "WhatsApp Monitor Client Loaded: "
                                + name
                );


            } catch (Exception e) {

                System.err.println(
                        "Failed to load WhatsApp client: "
                                + name
                );

                e.printStackTrace();
            }


            index++;
        }


        System.out.println(
                "WhatsApp Bot Monitor Clients Loaded: "
                        + index
        );
    }


    // =========================================================
    // CHECK ALL CLIENTS
    // =========================================================

    public List<WhatsAppBotMonitor> checkAllClients() {

        List<WhatsAppBotMonitor> clients =
                repository.findByActiveTrue();


        List<WhatsAppBotMonitor> results =
                new ArrayList<>();


        for (WhatsAppBotMonitor client : clients) {

            try {

                WhatsAppBotMonitor result =
                        checkClient(client);

                results.add(result);


            } catch (Exception e) {

                System.err.println(
                        "Error checking client: "
                                + client.getClientName()
                );


                client.setStatus(
                        WhatsAppBotMonitor.API_ERROR
                );


                client.setLastChecked(
                        LocalDateTime.now(INDIA_ZONE)
                );


                client.setMessage(
                        "Monitoring error"
                );


                results.add(
                        repository.save(client)
                );
            }
        }


        return results;
    }


    // =========================================================
    // CHECK SINGLE CLIENT
    // =========================================================

    public WhatsAppBotMonitor checkClient(
            WhatsAppBotMonitor client) {


        long startTime =
                System.currentTimeMillis();


        LocalDateTime checkTime =
                LocalDateTime.now(INDIA_ZONE);


        client.setLastChecked(checkTime);


        client.setStatus(
                WhatsAppBotMonitor.CHECKING
        );


        repository.save(client);


        try {

            // -------------------------------------------------
            // VALIDATE API URL
            // -------------------------------------------------

            if (client.getApiUrl() == null
                    || client.getApiUrl().isBlank()) {

                return updateError(
                        client,
                        WhatsAppBotMonitor.API_ERROR,
                        "API URL is missing",
                        null,
                        startTime
                );
            }


            // -------------------------------------------------
            // BUILD HTTP REQUEST
            // -------------------------------------------------

            HttpRequest.Builder requestBuilder =
                    HttpRequest.newBuilder()
                            .uri(
                                    URI.create(
                                            client.getApiUrl()
                                    )
                            )
                            .timeout(
                                    Duration.ofSeconds(30)
                            )
                            .header(
                                    "Accept",
                                    "application/json"
                            )
                            .header(
                                    "User-Agent",
                                    "DailyTaskManager"
                            );


            // -------------------------------------------------
            // ADD BEARER TOKEN
            // -------------------------------------------------

            if (client.getToken() != null
                    && !client.getToken().isBlank()) {


                String token =
                        client.getToken().trim();


                // Remove Bearer if it was already
                // added in properties

                if (token.regionMatches(
                        true,
                        0,
                        "Bearer ",
                        0,
                        7
                )) {

                    token =
                            token.substring(7).trim();
                }


                requestBuilder.header(
                        "Authorization",
                        "Bearer " + token
                );
            }


            HttpRequest request =
                    requestBuilder
                            .GET()
                            .build();


            // -------------------------------------------------
            // CALL API
            // -------------------------------------------------

            HttpResponse<String> response =
                    httpClient.send(
                            request,
                            HttpResponse.BodyHandlers.ofString()
                    );


            long responseTime =
                    System.currentTimeMillis()
                            - startTime;


            int statusCode =
                    response.statusCode();


            client.setStatusCode(statusCode);

            client.setResponseTime(responseTime);


            // -------------------------------------------------
            // AUTH ERROR
            // -------------------------------------------------

            if (statusCode == 401
                    || statusCode == 403) {

                return updateError(
                        client,
                        WhatsAppBotMonitor.AUTH_EXPIRED,
                        "Authentication expired or unauthorized",
                        responseTime,
                        startTime
                );
            }


            // -------------------------------------------------
            // API ERROR
            // -------------------------------------------------

            if (statusCode < 200
                    || statusCode >= 300) {

                return updateError(
                        client,
                        WhatsAppBotMonitor.API_ERROR,
                        "API returned HTTP " + statusCode,
                        responseTime,
                        startTime
                );
            }


            // -------------------------------------------------
            // EMPTY RESPONSE
            // -------------------------------------------------

            String responseBody =
                    response.body();


            if (responseBody == null
                    || responseBody.isBlank()) {

                return updateError(
                        client,
                        WhatsAppBotMonitor.API_ERROR,
                        "API returned empty response",
                        responseTime,
                        startTime
                );
            }


            // -------------------------------------------------
            // PARSE JSON
            // -------------------------------------------------

            JsonNode root =
                    objectMapper.readTree(
                            responseBody
                    );


            JsonNode resultArray =
                    root.path("data")
                            .path("result");


            // -------------------------------------------------
            // NO DATA
            // -------------------------------------------------

            if (!resultArray.isArray()
                    || resultArray.isEmpty()) {

                client.setStatus(
                        WhatsAppBotMonitor.NO_DATA
                );


                client.setMessage(
                        "No chat activity data available"
                );


                client.setInactiveMinutes(null);


                client.setLastChecked(checkTime);


                return repository.save(client);
            }


            // -------------------------------------------------
            // FIND LATEST updatedAt
            // -------------------------------------------------

            Instant latestActivity =
                    findLatestUpdatedAt(
                            resultArray
                    );


            if (latestActivity == null) {

                client.setStatus(
                        WhatsAppBotMonitor.NO_DATA
                );


                client.setMessage(
                        "No valid updatedAt found"
                );


                client.setInactiveMinutes(null);


                client.setLastChecked(checkTime);


                return repository.save(client);
            }


            // -------------------------------------------------
            // CALCULATE INACTIVITY
            // -------------------------------------------------

            Instant now =
                    Instant.now();


            long inactiveMinutes =
                    Duration.between(
                            latestActivity,
                            now
                    ).toMinutes();


            // Protect against future timestamps

            if (inactiveMinutes < 0) {

                inactiveMinutes = 0;
            }


            // -------------------------------------------------
            // CONVERT UTC TO IST
            // -------------------------------------------------

            LocalDateTime latestActivityIST =
                    latestActivity
                            .atZone(INDIA_ZONE)
                            .toLocalDateTime();


            client.setLastActivity(
                    latestActivityIST
            );


            client.setInactiveMinutes(
                    inactiveMinutes
            );


            // -------------------------------------------------
            // DETERMINE STATUS
            // -------------------------------------------------

            if (inactiveMinutes
                    < inactiveMinutesLimit) {


                client.setStatus(
                        WhatsAppBotMonitor.ACTIVE
                );


                client.setMessage(
                        "Bot activity detected"
                );


            } else {


                client.setStatus(
                        WhatsAppBotMonitor.NO_ACTIVITY
                );


                client.setMessage(
                        "No activity for "
                                + inactiveMinutes
                                + " minutes"
                );
            }


            client.setLastChecked(checkTime);


            return repository.save(client);


        } catch (DateTimeParseException e) {


            return updateError(
                    client,
                    WhatsAppBotMonitor.API_ERROR,
                    "Invalid updatedAt format",
                    null,
                    startTime
            );


        } catch (Exception e) {


            System.err.println(
                    "API check failed for "
                            + client.getClientName()
                            + ": "
                            + e.getClass()
                            .getSimpleName()
            );


            return updateError(
                    client,
                    WhatsAppBotMonitor.API_ERROR,
                    "Unable to connect to monitoring API",
                    null,
                    startTime
            );
        }
    }


    // =========================================================
    // FIND LATEST updatedAt
    // =========================================================

    private Instant findLatestUpdatedAt(
            JsonNode resultArray) {


        Instant latest = null;


        for (JsonNode item : resultArray) {


            String updatedAt =
                    item.path("updatedAt")
                            .asText(null);


            if (updatedAt == null
                    || updatedAt.isBlank()) {

                continue;
            }


            try {


                Instant current =
                        Instant.parse(updatedAt);


                if (latest == null
                        || current.isAfter(latest)) {

                    latest = current;
                }


            } catch (DateTimeParseException e) {


                // Skip invalid updatedAt

                System.err.println(
                        "Invalid updatedAt value skipped"
                );
            }
        }


        return latest;
    }


    // =========================================================
    // UPDATE ERROR
    // =========================================================

    private WhatsAppBotMonitor updateError(
            WhatsAppBotMonitor client,
            String status,
            String message,
            Long responseTime,
            long startTime) {


        client.setStatus(status);


        client.setMessage(message);


        client.setLastChecked(
                LocalDateTime.now(INDIA_ZONE)
        );


        if (responseTime != null) {

            client.setResponseTime(
                    responseTime
            );

        } else {

            client.setResponseTime(
                    System.currentTimeMillis()
                            - startTime
            );
        }


        return repository.save(client);
    }


    // =========================================================
    // GET ALL CLIENTS
    // =========================================================

    public List<WhatsAppBotMonitor> getAllClients() {

        return repository.findAll(
                Sort.by(
                        Sort.Direction.ASC,
                        "clientName"
                )
        );
    }


    // =========================================================
    // GET ACTIVE CLIENTS
    // =========================================================

    public List<WhatsAppBotMonitor> getActiveClients() {

        return repository.findByActiveTrue();
    }


    // =========================================================
    // GET SINGLE CLIENT
    // =========================================================

    public WhatsAppBotMonitor getClient(Long id) {

        return repository.findById(id)
                .orElseThrow(() ->
                        new RuntimeException(
                                "WhatsApp client not found: "
                                        + id
                        )
                );
    }


    // =========================================================
    // DELETE CLIENT
    // =========================================================

    public void deleteClient(Long id) {

        WhatsAppBotMonitor client =
                repository.findById(id)
                        .orElseThrow(() ->
                                new RuntimeException(
                                        "WhatsApp client not found: "
                                                + id
                                )
                        );


        repository.delete(client);
    }

}