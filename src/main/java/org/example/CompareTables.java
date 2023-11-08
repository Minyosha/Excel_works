package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

// Папка Excel
// All.xlsx - E - артикул, в F пишем да/нет
// Site.xlsx - L - артикул

//В папке "Excel" находятся файлы "All.xlsx" и "Site.xlsx". Открываем файл "All.xlsx". Если данные в ячейке "E" таблицы "All.xlsx" можно преобразовать в целое число,
//то ищем это число по всему столбцу "L" в файле "Site.xlsx". Если число из ячейки "E" в файле "All.xlsx" равно любому числу из столбца "L" в файле "Site.xlsx",
//то в ячейке "I" в таблице "All.xlsx" пишем "да" и пишем значение ячейки "E" в консоль. Если нет, то пишем "нет". Для сравнения данных необходимо будет преобразовать данные в столбцах "E" и "L" в целые числа.


public class CompareTables {
    public static void main(String[] args) {
        String allFilePath = "Excel/All.xlsx";
        String siteFilePath = "Excel/Site.xlsx";

        try (Workbook allWorkbook = new XSSFWorkbook(new FileInputStream(allFilePath));
             Workbook siteWorkbook = new XSSFWorkbook(new FileInputStream(siteFilePath))) {

            Sheet allSheet = allWorkbook.getSheetAt(0);
            Sheet siteSheet = siteWorkbook.getSheetAt(0);

            for (Row row : allSheet) {
                Cell valueCell = row.getCell(4); // Ячейка E

                if (valueCell != null) {
                    int numericValue = 0;
                    String valueType = "";

                    if (valueCell.getCellType() == CellType.NUMERIC) {
                        numericValue = (int) valueCell.getNumericCellValue();
                        valueType = "integer";
                    } else if (valueCell.getCellType() == CellType.STRING) {
                        numericValue = Integer.parseInt(valueCell.getStringCellValue());
                        valueType = "string";
                    }

                    for (Row siteRow : siteSheet) {
                        Cell siteValueCell = siteRow.getCell(11); // Ячейка L

                        if (siteValueCell != null) {
                            int siteValue = 0;
                            String siteValueType = "";

                            if (siteValueCell.getCellType() == CellType.NUMERIC) {
                                siteValue = (int) siteValueCell.getNumericCellValue();
                                siteValueType = "integer";
                            } else if (siteValueCell.getCellType() == CellType.STRING && isNumericCellValue(siteValueCell)) {
                                siteValue = Integer.parseInt(siteValueCell.getStringCellValue());
                                siteValueType = "string";
                            }

                            if (numericValue == siteValue) {
                                Cell resultCell = row.createCell(8); // Ячейка I
                                resultCell.setCellValue("да");

                                System.out.println("Значение " + valueCell + " (тип: " + valueType + ") в ячейке E таблицы All.xlsx совпадает с значением " + siteValueCell + " (тип: " + siteValueType + ") в столбце L таблицы Site.xlsx");
                                break;
                            }
                        }
                    }
                }
            }

            try (FileOutputStream outputStream = new FileOutputStream(allFilePath)) {
                allWorkbook.write(outputStream);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static boolean isNumericCellValue(Cell cell) {
        try {
            Double.parseDouble(cell.getStringCellValue());
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

}