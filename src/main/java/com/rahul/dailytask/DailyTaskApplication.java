
package com.rahul.dailytask;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableAsync;
import org.springframework.scheduling.annotation.EnableScheduling;

@SpringBootApplication
@EnableScheduling
@EnableAsync
public class DailyTaskApplication {

    public static void main(String[] args) {

        SpringApplication.run(DailyTaskApplication.class, args);

        System.out.println("====================================");
        System.out.println(" Daily Task Manager Started ");
        System.out.println("====================================");

        System.out.println(" URL Local : http://localhost:8080");
        System.out.println(
                " Dashboard Monitor: http://localhost:8080/dashboard-monitor.html");
        System.out.println(
                " API: http://localhost:8080/api/dashboard-monitor");

        System.out.println("------------------------------------");
        System.out.println(" General Scheduler: Every 5 minutes");
        System.out.println(" WhatsApp Bot Monitor: Enabled");
        System.out.println(" Inactivity Threshold: 15 minutes");

        System.out.println("------------------------------------");
        System.out.println(" WhatsApp Token Refresh Schedule (IST)");
        System.out.println(" 12:00 AM");
        System.out.println(" 06:00 AM");
        System.out.println(" 09:00 AM");
        System.out.println(" 10:00 AM");
        System.out.println(" 07:00 PM");
        System.out.println(" 10:00 PM");

        System.out.println("------------------------------------");
        System.out.println(
                " Tokens refresh only when authentication expires.");
        System.out.println(
                " Check All performs an immediate client check.");
        System.out.println("====================================");
    }
}
