package com.rahul.dailytask.scheduler;

import com.rahul.dailytask.entity.WhatsAppBotMonitor;
import com.rahul.dailytask.service.WhatsAppBotMonitorService;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
public class WhatsAppBotMonitorScheduler {

    private final WhatsAppBotMonitorService service;

    @Value("${whatsapp.monitor.enabled:true}")
    private boolean enabled;

    public WhatsAppBotMonitorScheduler(
            WhatsAppBotMonitorService service) {

        this.service = service;
    }

    // =========================================================
    // AUTOMATIC WHATSAPP BOT MONITORING
    // =========================================================

    @Scheduled(
            fixedDelayString =
                    "${whatsapp.monitor.fixed-delay:60000}"
    )
    public void monitorWhatsAppBots() {

        if (!enabled) {
            return;
        }

        System.out.println(
                "=========================================="
        );

        System.out.println(
                "Starting WhatsApp Bot Monitoring"
        );

        System.out.println(
                "=========================================="
        );

        try {

            List<WhatsAppBotMonitor> results =
                    service.checkAllClients();

            System.out.println(
                    "WhatsApp clients checked: "
                            + results.size()
            );

            for (WhatsAppBotMonitor client : results) {

                System.out.println(
                        "Client: "
                                + client.getClientName()
                                + " | Status: "
                                + client.getStatus()
                                + " | Inactive: "
                                + client.getInactiveMinutes()
                                + " min"
                );
            }

        } catch (Exception e) {

            System.err.println(
                    "WhatsApp Bot Monitoring failed"
            );

            e.printStackTrace();
        }

        System.out.println(
                "=========================================="
        );
    }
}