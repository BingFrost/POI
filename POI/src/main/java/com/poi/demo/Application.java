package com.poi.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication(scanBasePackages = "com.poi.demo")
public class Application {

    public static void main(String[] args) {
        SpringApplication.run(Application.class, args);
    }

}