package org.example;

public class Main {
    private static final String FILE_DIR = "C:\\Users\\vladm\\OneDrive\\Рабочий стол";

    public static void main(String[] args) {
        ExcelReader excelReader = new ExcelReader();
        DocxReader docxReader = new DocxReader();

        // Путь к исходным и новым файлам
        String excelInputPath = FILE_DIR + "\\b.xlsx";
        String excelOutputPath = FILE_DIR + "\\name.xlsx";

        String docxInputPath = FILE_DIR + "\\b.docx";
        String docxOutputPath = FILE_DIR + "\\updated_document.docx";

        // Поиск и замена в Excel
        String excelResult = excelReader.findAndReplace(excelInputPath, excelOutputPath, "проба", "дар");
        System.out.println("Excel Result:");
        System.out.println(excelResult);

        // Поиск и замена в DOCX
        String docxResult = docxReader.findAndReplace(docxInputPath, docxOutputPath, "пробник", "дар дар");
        System.out.println("DOCX Result:");
        System.out.println(docxResult);
    }
}
