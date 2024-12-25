package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Objects;

public class ExcelReader {

    public String findAndReplace(String excelFilePath, String outputFilePath, String replaceFrom, String replaceTo) {
        StringBuilder stringBuilder = new StringBuilder();

        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Чтение первой страницы

            for (Row row : sheet) {
                for (Cell cell : row) {
                    checkAndReplace(cell, replaceFrom, replaceTo);

                    switch (cell.getCellType()){
                        case NUMERIC -> {
                            stringBuilder.append(cell.getNumericCellValue());
                        }
                        case STRING -> {
                            stringBuilder.append(cell.getStringCellValue());
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }


        return stringBuilder.toString();
    }

    private void checkAndReplace(Cell cell, String replaceFrom, String replaceTo){
        if (Objects.requireNonNull(cell.getCellType()) == CellType.STRING) {
            if (cell.getStringCellValue().contains(replaceFrom)) {
                cell.setCellValue(cell.getStringCellValue().replaceAll(replaceFrom, replaceTo));
            }
        }
    }
}