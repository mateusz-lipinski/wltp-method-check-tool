package com.example.wltpcheck;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class WltpMethodCheckToolApplication implements CommandLineRunner {

    private final ExcelProcessingService excelProcessingService;

    public WltpMethodCheckToolApplication(ExcelProcessingService excelProcessingService) {
        this.excelProcessingService = excelProcessingService;
    }

    public static void main(String[] args) {
        SpringApplication.run(WltpMethodCheckToolApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
        if (args.length < 2) {
            System.err.println("Usage: wltp-method-check-tool <excel-file-path> <text-file-path>");
            return;
        }
        String excelFilePath = args[0];
        String textFilePath = args[1];
        excelProcessingService.process(excelFilePath, textFilePath);
    }
}
