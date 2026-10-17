
package com.rahul.dailytask.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.rahul.dailytask.entity.WhatsAppBotMonitor;

import org.springframework.core.env.Environment;
import org.springframework.stereotype.Service;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

@Service
public class TokenRefreshService {

    private final Environment environment;
    private final ObjectMapper objectMapper;
    private final HttpClient httpClient;

    public TokenRefreshService(
            Environment environment,
            ObjectMapper objectMapper) {

        this.environment = environment;
        this.objectMapper = objectMapper;

        this.httpClient = HttpClient.newBuilder()
                .connectTimeout(Duration.ofSeconds(15))
                .version(HttpClient.Version.HTTP_1_1)
                .build();
    }

    public String fetchNewToken(WhatsAppBotMonitor client) {

        if (client == null || client.getClientName() == null) {
            System.err.println("Token refresh skipped: client missing.");
            return null;
        }

        String prefix = findClientPrefix(client.getClientName());

        if (prefix == null) {
            System.err.println(
                    "No configuration found for " + client.getClientName());
            return null;
        }

        String authUrl = getProperty(prefix + ".auth-url");
        String email = getProperty(prefix + ".auth-email");
        String password = getProperty(prefix + ".auth-password");
        String basicAuth = getProperty(prefix + ".auth-basic");

        if (isBlank(authUrl) || isBlank(email) || isBlank(password)
                || "CHANGE_ME".equalsIgnoreCase(password)) {

            System.err.println(
                    "Login configuration missing or using placeholder for "
                            + client.getClientName());
            return null;
        }

        try {
            String requestBody = objectMapper.createObjectNode()
                    .put("email", email)
                    .put("password", password)
                    .toString();

            HttpRequest.Builder builder = HttpRequest.newBuilder()
                    .uri(URI.create(authUrl))
                    .timeout(Duration.ofSeconds(30))
                    .header("Accept", "application/json")
                    .header("Content-Type", "application/json")
                    .header("User-Agent", "DailyTaskManager");

            if (!isBlank(basicAuth)
                    && !"CHANGE_ME".equalsIgnoreCase(basicAuth)) {

                String value = basicAuth.trim();

                if (value.regionMatches(true, 0, "Basic ", 0, 6)) {
                    value = value.substring(6).trim();
                }

                builder.header("Authorization", "Basic " + value);
            }

            HttpRequest request = builder
                    .POST(HttpRequest.BodyPublishers.ofString(
                            requestBody, StandardCharsets.UTF_8))
                    .build();

            HttpResponse<String> response = httpClient.send(
                    request,
                    HttpResponse.BodyHandlers.ofString(
                            StandardCharsets.UTF_8));

            System.out.println("Login API HTTP status for "
                    + client.getClientName() + ": "
                    + response.statusCode());

            if (response.statusCode() < 200
                    || response.statusCode() >= 300) {

                System.err.println("Login API failed for "
                        + client.getClientName()
                        + ". HTTP " + response.statusCode());

                return null;
            }

            String body = response.body();

            if (isBlank(body)) {
                System.err.println("Login API returned an empty response for "
                        + client.getClientName());
                return null;
            }

            JsonNode root = objectMapper.readTree(body);
            JsonNode data = root.path("data");

            // Diagnostic logs: field names only, never token values.
            System.out.println("Login response root fields for "
                    + client.getClientName() + ": " + fieldNames(root));

            System.out.println("Login response data fields for "
                    + client.getClientName() + ": " + fieldNames(data));

            System.out.println("Token field checks for "
                    + client.getClientName()
                    + ": data.authToken=" + hasText(data, "authToken")
                    + ", data.token=" + hasText(data, "token")
                    + ", data.accessToken=" + hasText(data, "accessToken")
                    + ", root.authToken=" + hasText(root, "authToken")
                    + ", root.token=" + hasText(root, "token")
                    + ", root.accessToken=" + hasText(root, "accessToken"));

            // Try common token field names.
            String[] tokenFields = {
                    "authToken",
                    "token",
                    "accessToken"
            };

            JsonNode tokenNode = findTokenNode(data, tokenFields);

            if (tokenNode == null) {
                tokenNode = findTokenNode(root, tokenFields);
            }

            if (tokenNode == null) {
                System.err.println(
                        "Login returned HTTP success, but no supported "
                                + "token field was found for "
                                + client.getClientName());
                return null;
            }

            String token = normalizeToken(tokenNode.asText());

            if (isBlank(token)) {
                System.err.println("Extracted token is empty for "
                        + client.getClientName());
                return null;
            }

            System.out.println("Token received for "
                    + client.getClientName()
                    + ". Value not logged.");

            return token;

        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            System.err.println("Login request interrupted for "
                    + client.getClientName());
            return null;

        } catch (Exception e) {
            System.err.println("Token refresh failed for "
                    + client.getClientName()
                    + ": " + e.getClass().getSimpleName());
            return null;
        }
    }

    private JsonNode findTokenNode(
            JsonNode parent,
            String[] fields) {

        if (parent == null || !parent.isObject()) {
            return null;
        }

        for (String field : fields) {
            JsonNode node = parent.path(field);

            if (node.isTextual() && !isBlank(node.asText())) {
                return node;
            }
        }

        return null;
    }

    private String fieldNames(JsonNode node) {

        if (node == null || !node.isObject()) {
            return "not-an-object";
        }

        List<String> names = new ArrayList<>();
        node.fieldNames().forEachRemaining(names::add);

        return names.toString();
    }

    private boolean hasText(JsonNode node, String field) {

        JsonNode value = node.path(field);

        return value.isTextual() && !value.asText().isBlank();
    }

    private String findClientPrefix(String clientName) {

        for (int index = 0; index < 100; index++) {

            String prefix = "monitor.clients[" + index + "]";
            String configuredName = getProperty(prefix + ".name");

            if (isBlank(configuredName)) {
                continue;
            }

            if (configuredName.equalsIgnoreCase(clientName.trim())) {
                return prefix;
            }
        }

        return null;
    }

    private String getProperty(String key) {

        String value = environment.getProperty(key);

        return value == null ? null : value.trim();
    }

    private String normalizeToken(String token) {

        if (token == null) {
            return null;
        }

        String normalized = token.trim();

        if (normalized.regionMatches(true, 0, "Bearer ", 0, 7)) {
            normalized = normalized.substring(7).trim();
        }

        return normalized;
    }

    private boolean isBlank(String value) {

        return value == null || value.isBlank();
    }
}
