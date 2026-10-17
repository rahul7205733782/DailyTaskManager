
package com.rahul.dailytask.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;

@Component
@ConfigurationProperties(prefix = "monitor")
public class WhatsAppMonitorProperties {

    private List<Client> clients = new ArrayList<>();

    public List<Client> getClients() {
        return clients;
    }

    public void setClients(List<Client> clients) {
        this.clients = clients;
    }

    public static class Client {

        private String name;
        private String apiUrl;
        private String token;
        private String authUrl;
        private String authEmail;
        private String authPassword;
        private String authBasic;

        // Client name
        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        // All Chat API URL
        public String getApiUrl() {
            return apiUrl;
        }

        public void setApiUrl(String apiUrl) {
            this.apiUrl = apiUrl;
        }

        // Authentication token
        public String getToken() {
            return token;
        }

        public void setToken(String token) {
            this.token = token;
        }

        // Login API URL
        public String getAuthUrl() {
            return authUrl;
        }

        public void setAuthUrl(String authUrl) {
            this.authUrl = authUrl;
        }

        // Login email
        public String getAuthEmail() {
            return authEmail;
        }

        public void setAuthEmail(String authEmail) {
            this.authEmail = authEmail;
        }

        // Login password
        public String getAuthPassword() {
            return authPassword;
        }

        public void setAuthPassword(String authPassword) {
            this.authPassword = authPassword;
        }

        // Basic Auth credentials (Base64 value)
        public String getAuthBasic() {
            return authBasic;
        }

        public void setAuthBasic(String authBasic) {
            this.authBasic = authBasic;
        }
    }
}
