package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class DocxReader {

    public String findAndReplace(String docxFilePath, String outputFilePath, String replaceFrom, String replaceTo) {
        StringBuilder stringBuilder = new StringBuilder();

        try (FileInputStream fis = new FileInputStream(docxFilePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // Обрабатываем все параграфы
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (XWPFParagraph paragraph : paragraphs) {
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null) {
                    for (XWPFRun run : runs) {
                        String text = run.getText(0); // Получаем текст текущего Run
                        if (text != null && text.contains(replaceFrom)) {
                            text = text.replace(replaceFrom, replaceTo);
                            run.setText(text, 0); // Обновляем текст в текущем Run
                        }
                    }
                }
                stringBuilder.append(paragraph.getText()).append("\n");
            }

            // Обрабатываем таблицы (если есть)
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            List<XWPFRun> runs = paragraph.getRuns();
                            if (runs != null) {
                                for (XWPFRun run : runs) {
                                    String text = run.getText(0);
                                    if (text != null && text.contains(replaceFrom)) {
                                        text = text.replace(replaceFrom, replaceTo);
                                        run.setText(text, 0);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Сохраняем изменения в новый файл
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                document.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return stringBuilder.toString();
    }
}
