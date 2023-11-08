package org.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

public class CopyBase {
    public static void main(String[] args) {
        String baseFilePath = "Excel/Base table.xlsx";
        String allFilePath = "Excel/All.xlsx";

        try {
            copyFile(baseFilePath, allFilePath);
            System.out.println("Копирование файла успешно завершено.");
        } catch (IOException e) {
            System.out.println("Ошибка при копировании файла: " + e.getMessage());
        }
    }

    private static void copyFile(String sourceFilePath, String destinationFilePath) throws IOException {
        Path sourcePath = Path.of(sourceFilePath);
        Path destinationPath = Path.of(destinationFilePath);

        // Создаем копию файла в директории назначения
        Files.copy(sourcePath, destinationPath, StandardCopyOption.REPLACE_EXISTING);
    }
}
