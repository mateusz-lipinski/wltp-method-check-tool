package com.example.wltpcheck;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

@SpringBootTest
class ExcelProcessingServiceTest {

    @Autowired
    private ExcelProcessingService service;

    @TempDir
    Path tempDir;

    @Test
    void testIsValidUpperCaseName_valid() {
        assertThat(service.isValidUpperCaseName("EXHAUST_EMISSION_TEST_1_BIV_CO")).isTrue();
        assertThat(service.isValidUpperCaseName("SIMPLE")).isTrue();
        assertThat(service.isValidUpperCaseName("WITH_123_NUMBERS")).isTrue();
    }

    @Test
    void testIsValidUpperCaseName_invalid() {
        assertThat(service.isValidUpperCaseName("lower_case")).isFalse();
        assertThat(service.isValidUpperCaseName("Mixed_Case")).isFalse();
        assertThat(service.isValidUpperCaseName("HAS SPACE")).isFalse();
        assertThat(service.isValidUpperCaseName("")).isFalse();
        assertThat(service.isValidUpperCaseName(null)).isFalse();
    }

    @Test
    void testToCamelCase() {
        assertThat(service.toCamelCase("EXHAUST_EMISSION_TEST_1_BIV_CO")).isEqualTo("exhaustEmissionTest1BivCo");
        assertThat(service.toCamelCase("SIMPLE")).isEqualTo("simple");
        assertThat(service.toCamelCase("HELLO_WORLD")).isEqualTo("helloWorld");
        assertThat(service.toCamelCase("TEST_123")).isEqualTo("test123");
    }

    @Test
    void testProcess_writesDefaultStyleWhenFound(@TempDir Path dir) throws IOException {
        Path excelFile = dir.resolve("test.xlsx");
        Path textFile = dir.resolve("test.txt");

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            // rows 0-2 are skipped (rows 1-3 in 1-based), row 3 (index) is first processed
            for (int i = 0; i < 3; i++) {
                sheet.createRow(i);
            }
            Row row = sheet.createRow(3);
            row.createCell(0).setCellValue("EXHAUST_EMISSION_TEST_1_BIV_CO");
            try (FileOutputStream fos = new FileOutputStream(excelFile.toFile())) {
                workbook.write(fos);
            }
        }

        Files.writeString(textFile, "private String exhaustEmissionTest1BivCo;\n");

        service.process(excelFile.toString(), textFile.toString());

        try (Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(excelFile.toFile())) {
            Sheet sheet = workbook.getSheetAt(0);
            Cell cell = sheet.getRow(3).getCell(1);
            assertThat(cell.getStringCellValue()).isEqualTo("exhaustEmissionTest1BivCo");
            // Default font has no color set (color index 0 or automatic)
            short colorIndex = workbook.getFontAt(cell.getCellStyle().getFontIndex()).getColor();
            assertThat(colorIndex).isNotEqualTo(org.apache.poi.ss.usermodel.IndexedColors.RED.getIndex());
        }
    }

    @Test
    void testProcess_writesRedStyleWhenNotFound(@TempDir Path dir) throws IOException {
        Path excelFile = dir.resolve("test.xlsx");
        Path textFile = dir.resolve("test.txt");

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            for (int i = 0; i < 3; i++) {
                sheet.createRow(i);
            }
            Row row = sheet.createRow(3);
            row.createCell(0).setCellValue("EXHAUST_EMISSION_TEST_1_BIV_CO");
            try (FileOutputStream fos = new FileOutputStream(excelFile.toFile())) {
                workbook.write(fos);
            }
        }

        Files.writeString(textFile, "// no matching field here\n");

        service.process(excelFile.toString(), textFile.toString());

        try (Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(excelFile.toFile())) {
            Sheet sheet = workbook.getSheetAt(0);
            Cell cell = sheet.getRow(3).getCell(1);
            assertThat(cell.getStringCellValue()).isEqualTo("exhaustEmissionTest1BivCo");
            short colorIndex = workbook.getFontAt(cell.getCellStyle().getFontIndex()).getColor();
            assertThat(colorIndex).isEqualTo(org.apache.poi.ss.usermodel.IndexedColors.RED.getIndex());
        }
    }

    @Test
    void testProcess_skipsInvalidRows(@TempDir Path dir) throws IOException {
        Path excelFile = dir.resolve("test.xlsx");
        Path textFile = dir.resolve("test.txt");

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet();
            for (int i = 0; i < 3; i++) {
                sheet.createRow(i);
            }
            Row row = sheet.createRow(3);
            row.createCell(0).setCellValue("not_valid_UPPER");
            try (FileOutputStream fos = new FileOutputStream(excelFile.toFile())) {
                workbook.write(fos);
            }
        }

        Files.writeString(textFile, "");

        service.process(excelFile.toString(), textFile.toString());

        try (Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(excelFile.toFile())) {
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(3);
            Cell cell = row.getCell(1);
            assertThat(cell).isNull();
        }
    }
}
