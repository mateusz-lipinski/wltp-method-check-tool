package com.example.wltpcheck;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.regex.Pattern;

@Service
public class ExcelProcessingService {

    private static final Pattern VALID_NAME_PATTERN = Pattern.compile("^[A-Z0-9_]+$");
    private static final int FIRST_DATA_ROW_INDEX = 3; // 0-based: row 4 is index 3

    public void process(String excelFilePath, String textFilePath) throws IOException {
        String textFileContent = Files.readString(Paths.get(textFilePath), StandardCharsets.UTF_8);

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            CellStyle defaultStyle = workbook.createCellStyle();
            Font defaultFont = workbook.createFont();
            defaultStyle.setFont(defaultFont);

            CellStyle redStyle = workbook.createCellStyle();
            Font redFont = workbook.createFont();
            redFont.setColor(IndexedColors.RED.getIndex());
            redStyle.setFont(redFont);

            for (int rowIndex = FIRST_DATA_ROW_INDEX; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }

                Cell firstCell = row.getCell(0);
                if (firstCell == null || firstCell.getCellType() != CellType.STRING) {
                    continue;
                }

                String cellValue = firstCell.getStringCellValue().trim();
                if (!isValidUpperCaseName(cellValue)) {
                    continue;
                }

                String variableName = toCamelCase(cellValue);
                String searchText = "private String " + variableName + ";";

                Cell secondCell = row.getCell(1);
                if (secondCell == null) {
                    secondCell = row.createCell(1);
                }

                secondCell.setCellValue(variableName);
                if (textFileContent.contains(searchText)) {
                    secondCell.setCellStyle(defaultStyle);
                } else {
                    secondCell.setCellStyle(redStyle);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                workbook.write(fos);
            }
        }
    }

    boolean isValidUpperCaseName(String value) {
        if (value == null || value.isEmpty()) {
            return false;
        }
        return VALID_NAME_PATTERN.matcher(value).matches();
    }

    String toCamelCase(String upperCaseName) {
        String[] parts = upperCaseName.split("_");
        StringBuilder result = new StringBuilder();
        for (int i = 0; i < parts.length; i++) {
            String part = parts[i];
            if (part.isEmpty()) {
                continue;
            }
            if (i == 0) {
                result.append(part.toLowerCase());
            } else {
                if (Character.isLetter(part.charAt(0))) {
                    result.append(Character.toUpperCase(part.charAt(0)));
                    result.append(part.substring(1).toLowerCase());
                } else {
                    result.append(part.toLowerCase());
                }
            }
        }
        return result.toString();
    }
}
