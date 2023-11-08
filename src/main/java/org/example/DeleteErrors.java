package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class DeleteErrors {
    public static void main(String[] args) {
        String allFilePath = "Excel/All.xlsx";

        try (Workbook allWorkbook = new XSSFWorkbook(new FileInputStream(allFilePath))) {
            Sheet allSheet = allWorkbook.getSheetAt(0);
            deleteInvalidValues(allSheet);

            try (FileOutputStream outputStream = new FileOutputStream(allFilePath)) {
                allWorkbook.write(outputStream);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void deleteInvalidValues(Sheet sheet) {
        DataFormatter dataFormatter = new DataFormatter();
        for (Row row : sheet) {
            Cell valueCell = row.getCell(4); // Ячейка E
            Cell resultCell = row.getCell(8); // Ячейка I

            if (valueCell != null && resultCell != null) {
                String value = dataFormatter.formatCellValue(valueCell);
                String result = dataFormatter.formatCellValue(resultCell);

                System.out.println("Значение ячейки E: " + value);
                System.out.println("Значение ячейки I до изменения: " + result);

                if (!isCorrectValue(value)) {
                    resultCell.setCellValue(""); // Очистить значение ячейки "I"
                    System.out.println("Удалено значение из ячейки I");
                } else {
                    System.out.println("Оставлено значение в ячейке I");
                }

                result = dataFormatter.formatCellValue(resultCell);
                System.out.println("Значение ячейки I после изменения: " + result);
                System.out.println();
            }
        }
    }

    private static boolean isCorrectValue(String value) {
        // Проверить, является ли значение правильным
        // Здесь вы можете добавить свою логику проверки
        // Например, проверка на числовое значение или другие критерии
        // Возвращайте true, если значение является правильным, и false в противном случае
        return isNumeric(value);
    }

    private static boolean isNumeric(String value) {
        try {
            Double.parseDouble(value);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
}