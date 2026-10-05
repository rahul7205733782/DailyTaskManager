package com.rahul.dailytask.controller;

import com.rahul.dailytask.entity.WhatsAppBotMonitor;
import com.rahul.dailytask.repository.WhatsAppBotMonitorRepository;
import com.rahul.dailytask.service.WhatsAppBotMonitorService;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/whatsapp-monitor")
@CrossOrigin
public class WhatsAppBotMonitorController {

    private final WhatsAppBotMonitorService service;
    private final WhatsAppBotMonitorRepository repository;

    // =========================================================
    // CONSTRUCTOR
    // =========================================================

    public WhatsAppBotMonitorController(
            WhatsAppBotMonitorService service,
            WhatsAppBotMonitorRepository repository) {

        this.service = service;
        this.repository = repository;
    }


    // =========================================================
    // GET ALL CLIENTS
    // =========================================================

    @GetMapping("/clients")
    public ResponseEntity<List<WhatsAppBotMonitor>> getAllClients() {

        return ResponseEntity.ok(
                repository.findAll()
        );
    }


    // =========================================================
    // ADD NEW CLIENT
    // =========================================================

    @PostMapping("/clients")
    public ResponseEntity<WhatsAppBotMonitor> addClient(
            @RequestBody WhatsAppBotMonitor client) {

        // Make sure this is treated as a new client
        client.setId(null);

        // Initial status
        client.setStatus(
                WhatsAppBotMonitor.CHECKING
        );

        client.setActive(true);

        // Clear monitoring result fields
        client.setLastActivity(null);
        client.setLastChecked(null);
        client.setInactiveMinutes(null);
        client.setStatusCode(null);
        client.setResponseTime(null);
        client.setMessage(null);

        // Save client
        WhatsAppBotMonitor savedClient =
                repository.save(client);

        return ResponseEntity.ok(savedClient);
    }


    // =========================================================
    // CHECK ONE CLIENT
    // =========================================================

    @PostMapping("/check/{id}")
    public ResponseEntity<WhatsAppBotMonitor> checkClient(
            @PathVariable Long id) {

        /*
         * The service currently expects:
         *
         * checkClient(WhatsAppBotMonitor client)
         *
         * but the URL gives us:
         *
         * Long id
         *
         * So first find the client from DB.
         */

        WhatsAppBotMonitor client =
                repository.findById(id)
                        .orElseThrow(() ->
                                new RuntimeException(
                                        "WhatsApp client not found with id: " + id
                                )
                        );

        // Send the actual entity to the service
        WhatsAppBotMonitor result =
                service.checkClient(client);

        return ResponseEntity.ok(result);
    }


    // =========================================================
    // CHECK ALL CLIENTS
    // =========================================================

    @PostMapping("/check-all")
    public ResponseEntity<List<WhatsAppBotMonitor>> checkAllClients() {

        List<WhatsAppBotMonitor> results =
                service.checkAllClients();

        return ResponseEntity.ok(results);
    }


    // =========================================================
    // DELETE CLIENT
    // =========================================================

    @DeleteMapping("/clients/{id}")
    public ResponseEntity<Void> deleteClient(
            @PathVariable Long id) {

        if (!repository.existsById(id)) {
            return ResponseEntity.notFound().build();
        }

        repository.deleteById(id);

        return ResponseEntity.noContent().build();
    }
}