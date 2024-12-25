package org.example;

import javafx.application.Application;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.VBox;

import java.io.File;


public class MainApp extends Application {

    @Override
    public void start(Stage primaryStage) {
        // Элементы интерфейса
        Label label = new Label("Выберите файл и замените текст:");
        TextField findField = new TextField();
        findField.setPromptText("Что искать?");
        TextField replaceField = new TextField();
        replaceField.setPromptText("На что заменить?");
        Button selectFileButton = new Button("Выбрать файл");
        Button processButton = new Button("Обработать");
        Label resultLabel = new Label();

        FileChooser fileChooser = new FileChooser();
        final File[] selectedFile = new File[1];

        selectFileButton.setOnAction(e -> {
            File file = fileChooser.showOpenDialog(primaryStage);
            if (file != null) {
                selectedFile[0] = file;
                label.setText("Выбран файл: " + file.getName());
            }
        });

        processButton.setOnAction(e -> {
            if (selectedFile[0] != null && !findField.getText().isEmpty() && !replaceField.getText().isEmpty()) {
                String filePath = selectedFile[0].getAbsolutePath();
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
                resultLabel.setText("Пожалуйста, заполните все поля и выберите файл.");
            }
        });

        // Размещение элементов на экране
        VBox vBox = new VBox(10, label, findField, replaceField, selectFileButton, processButton, resultLabel);
        Scene scene = new Scene(vBox, 400, 300);

        primaryStage.setTitle("Поиск и замена в файлах");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch();
    }
}
