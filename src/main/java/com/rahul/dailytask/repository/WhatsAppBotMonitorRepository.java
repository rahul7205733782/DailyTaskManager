package com.rahul.dailytask.repository;

import com.rahul.dailytask.entity.WhatsAppBotMonitor;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface WhatsAppBotMonitorRepository
        extends JpaRepository<WhatsAppBotMonitor, Long> {

    // Get all active WhatsApp bot clients
    List<WhatsAppBotMonitor> findByActiveTrue();

    // Find client by name
    Optional<WhatsAppBotMonitor> findByClientNameIgnoreCase(String clientName);

    // Search clients by name
    List<WhatsAppBotMonitor> findByClientNameContainingIgnoreCase(String clientName);

    // Get clients by monitoring status
    List<WhatsAppBotMonitor> findByStatusIgnoreCase(String status);

    // Count clients by status
    long countByStatusIgnoreCase(String status);

    // Count active clients
    long countByActiveTrue();
}