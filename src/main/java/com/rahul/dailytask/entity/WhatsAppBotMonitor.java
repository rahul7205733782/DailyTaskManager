package com.rahul.dailytask.entity;

import com.fasterxml.jackson.annotation.JsonProperty;

import jakarta.persistence.*;

import java.time.LocalDateTime;

@Entity
@Table(name = "whatsapp_bot_monitor")
public class WhatsAppBotMonitor {

    // =========================================================
    // STATUS CONSTANTS
    // =========================================================

    public static final String ACTIVE = "ACTIVE";

    public static final String NO_ACTIVITY = "NO_ACTIVITY";

    public static final String API_ERROR = "API_ERROR";

    public static final String AUTH_EXPIRED = "AUTH_EXPIRED";

    public static final String NO_DATA = "NO_DATA";

    public static final String CHECKING = "CHECKING";


    // =========================================================
    // PRIMARY KEY
    // =========================================================

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;


    // =========================================================
    // CLIENT INFORMATION
    // =========================================================

    @Column(nullable = false)
    private String clientName;

    @Column(nullable = false, length = 1000)
    private String apiUrl;


    // =========================================================
    // AUTHENTICATION TOKEN
    // =========================================================
    // WRITE_ONLY means:
    //
    // 1. Dashboard can send the token in POST request
    // 2. Backend can read and save the token
    // 3. Monitoring service can use getToken()
    // 4. Token will NOT be returned in GET API JSON response
    //
    // This prevents the JWT from being exposed to the frontend.

    @JsonProperty(access = JsonProperty.Access.WRITE_ONLY)
    @Column(length = 2000)
    private String token;


    // =========================================================
    // MONITORING INFORMATION
    // =========================================================

    @Column(nullable = false)
    private String status;

    private LocalDateTime lastActivity;

    private LocalDateTime lastChecked;

    private Long inactiveMinutes;

    private Integer statusCode;

    private Long responseTime;

    @Column(length = 500)
    private String message;

    private Boolean active = true;


    // =========================================================
    // DEFAULT CONSTRUCTOR
    // =========================================================

    public WhatsAppBotMonitor() {
    }


    // =========================================================
    // PARAMETERIZED CONSTRUCTOR
    // =========================================================

    public WhatsAppBotMonitor(
            String clientName,
            String apiUrl,
            String token) {

        this.clientName = clientName;
        this.apiUrl = apiUrl;
        this.token = token;
        this.status = CHECKING;
        this.active = true;
    }


    // =========================================================
    // GET ID
    // =========================================================

    public Long getId() {
        return id;
    }


    // =========================================================
    // SET ID
    // =========================================================

    public void setId(Long id) {
        this.id = id;
    }


    // =========================================================
    // GET CLIENT NAME
    // =========================================================

    public String getClientName() {
        return clientName;
    }


    // =========================================================
    // SET CLIENT NAME
    // =========================================================

    public void setClientName(String clientName) {
        this.clientName = clientName;
    }


    // =========================================================
    // GET API URL
    // =========================================================

    public String getApiUrl() {
        return apiUrl;
    }


    // =========================================================
    // SET API URL
    // =========================================================

    public void setApiUrl(String apiUrl) {
        this.apiUrl = apiUrl;
    }


    // =========================================================
    // GET TOKEN
    // =========================================================
    // Used internally by WhatsAppBotMonitorService.
    //
    // The token is WRITE_ONLY for JSON, so it will not be
    // included when this entity is returned by REST APIs.

    public String getToken() {
        return token;
    }


    // =========================================================
    // SET TOKEN
    // =========================================================

    public void setToken(String token) {
        this.token = token;
    }


    // =========================================================
    // GET STATUS
    // =========================================================

    public String getStatus() {
        return status;
    }


    // =========================================================
    // SET STATUS
    // =========================================================

    public void setStatus(String status) {
        this.status = status;
    }


    // =========================================================
    // GET LAST ACTIVITY
    // =========================================================

    public LocalDateTime getLastActivity() {
        return lastActivity;
    }


    // =========================================================
    // SET LAST ACTIVITY
    // =========================================================

    public void setLastActivity(
            LocalDateTime lastActivity) {

        this.lastActivity = lastActivity;
    }


    // =========================================================
    // GET LAST CHECKED
    // =========================================================

    public LocalDateTime getLastChecked() {
        return lastChecked;
    }


    // =========================================================
    // SET LAST CHECKED
    // =========================================================

    public void setLastChecked(
            LocalDateTime lastChecked) {

        this.lastChecked = lastChecked;
    }


    // =========================================================
    // GET INACTIVE MINUTES
    // =========================================================

    public Long getInactiveMinutes() {
        return inactiveMinutes;
    }


    // =========================================================
    // SET INACTIVE MINUTES
    // =========================================================

    public void setInactiveMinutes(
            Long inactiveMinutes) {

        this.inactiveMinutes = inactiveMinutes;
    }


    // =========================================================
    // GET STATUS CODE
    // =========================================================

    public Integer getStatusCode() {
        return statusCode;
    }


    // =========================================================
    // SET STATUS CODE
    // =========================================================

    public void setStatusCode(
            Integer statusCode) {

        this.statusCode = statusCode;
    }


    // =========================================================
    // GET RESPONSE TIME
    // =========================================================

    public Long getResponseTime() {
        return responseTime;
    }


    // =========================================================
    // SET RESPONSE TIME
    // =========================================================

    public void setResponseTime(
            Long responseTime) {

        this.responseTime = responseTime;
    }


    // =========================================================
    // GET MESSAGE
    // =========================================================

    public String getMessage() {
        return message;
    }


    // =========================================================
    // SET MESSAGE
    // =========================================================

    public void setMessage(String message) {
        this.message = message;
    }


    // =========================================================
    // GET ACTIVE
    // =========================================================

    public Boolean getActive() {
        return active;
    }


    // =========================================================
    // SET ACTIVE
    // =========================================================

    public void setActive(Boolean active) {
        this.active = active;
    }


    // =========================================================
    // HELPER METHODS
    // =========================================================

    public boolean isActiveStatus() {

        return ACTIVE.equalsIgnoreCase(status);
    }


    public boolean isNoActivity() {

        return NO_ACTIVITY.equalsIgnoreCase(status);
    }


    public boolean isApiError() {

        return API_ERROR.equalsIgnoreCase(status);
    }


    public boolean isAuthExpired() {

        return AUTH_EXPIRED.equalsIgnoreCase(status);
    }


    public boolean isNoData() {

        return NO_DATA.equalsIgnoreCase(status);
    }


    public boolean isChecking() {

        return CHECKING.equalsIgnoreCase(status);
    }
}