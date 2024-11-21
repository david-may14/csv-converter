package com.personal;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class XlsxToCsvConverter {

    public static void convertXlsxToCsv(String inputFilePath, String outputFilePath) throws IOException {
        // Abre el archivo XLSX
        try (Workbook workbook = new XSSFWorkbook(Files.newInputStream(Paths.get(inputFilePath)))) {
            // Obtén la primera hoja
            Sheet sheet = workbook.getSheetAt(0);

            // Prepara el archivo CSV de salida
            try (BufferedWriter writer = Files.newBufferedWriter(Paths.get(outputFilePath))) {
                for (Row row : sheet) {
                    StringBuilder csvRow = new StringBuilder();
                    for (Cell cell : row) {
                        String cellValue = getCellValueAsString(cell);
                        csvRow.append(cellValue).append(","); // Usa ',' como separador
                    }
                    // Elimina la última coma y escribe la línea
                    if (csvRow.length() > 0) {
                        csvRow.setLength(csvRow.length() - 1);
                    }
                    writer.write(csvRow.toString());
                    writer.newLine();
                }
            }
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "";
        }
    }

    public static void main(String[] args) {
        String inputFilePath = "ruta/del/archivo.xlsx";
        String outputFilePath = "ruta/del/archivo.csv";
        try {
            convertXlsxToCsv(inputFilePath, outputFilePath);
            System.out.println("Conversión completada: " + outputFilePath);
        } catch (IOException e) {
            System.err.println("Error durante la conversión: " + e.getMessage());
        }
    }
}
