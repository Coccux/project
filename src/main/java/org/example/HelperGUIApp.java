package org.example;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class HelperGUIApp extends Application {

    public static void main(String[] args) {
        launch();
    }

    @Override
    public void start(Stage primaryStage) {
        Label label = new Label("Выберите директорию с файлами для замены текста:");
        TextField findField = new TextField();
        findField.setPromptText("Что искать?");
        TextField replaceField = new TextField();
        replaceField.setPromptText("На что заменять?");
        Button selectFileButton = new Button("Выбрать директорию");
        Button processButton = new Button("Найти и заменить");
        Label resultLabel = new Label();

        DirectoryChooser directoryChooser = new DirectoryChooser();
        final List<File> selectedFiles = new ArrayList<>();

        selectFileButton.setOnAction(e -> {
            File file = directoryChooser.showDialog(primaryStage);
            if (file != null) {
                selectedFiles.addAll(readFiles(file.toPath()));
            }
        });

        processButton.setOnAction(e -> {
            selectedFiles.parallelStream()
                    .forEach(file -> {
                        if (!findField.getText().isEmpty() && !replaceField.getText().isEmpty()) {
                            String filePath = file.getAbsolutePath();
                            String outputPath = filePath.substring(0, filePath.lastIndexOf(".")) + "_updated" + filePath.substring(filePath.lastIndexOf("."));

                            try {
                                if (filePath.endsWith(".xlsx")) {
                                    ExcelReader excelReader = new ExcelReader();
                                    excelReader.findAndReplace(filePath, outputPath, findField.getText(), replaceField.getText());
                                    resultLabel.setText("Excel файл обработан: " + outputPath);
                                } else if (filePath.endsWith(".docx")) {
                                    DocxReader docxReader = new DocxReader();
                                    docxReader.findAndReplace(filePath, outputPath, findField.getText(), replaceField.getText());
                                    resultLabel.setText("DOCX файл обработан: " + outputPath);
                                } else {
                                    resultLabel.setText("Неподдерживаемый формат файла!");
                                }
                            } catch (Exception ex) {
                                ex.printStackTrace();
                                resultLabel.setText("Произошла ошибка!");
                            }
                        } else {
                            resultLabel.setText("Пожалуйста, заполните все поля и выберите директорию.");
                        }
                    });
        });

        VBox vBox = new VBox(10, label, findField, replaceField, selectFileButton, processButton, resultLabel);
        Scene scene = new Scene(vBox, 400, 300);

        primaryStage.setTitle("Поиск и замена в файлах");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private List<File> readFiles(Path fromPath) {
        try {
            return Files.walk(fromPath)
                    .filter(path -> path.toString().endsWith(".xlsx") || path.toString().endsWith(".docx"))
                    .map(Path::toFile)
                    .toList();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
