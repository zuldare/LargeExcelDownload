package com.jh.largeexceldownload.excel.streaming;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.IOException;

public class LargeExcelFromTemplate {

    public void generate(OutputStream outputStream) {
        String templatePath = "path_to_your_template/template.xlsx"; // Ruta a la plantilla XLSX
        String outputPath = "path_to_output/output.xlsx"; // Ruta al archivo XLSX de salida

        try (InputStream is = new FileInputStream(templatePath)) {
            XSSFWorkbook workbookTemplate = new XSSFWorkbook(is);
            SXSSFWorkbook workbook = new SXSSFWorkbook(workbookTemplate, 100);

            Sheet sheet = workbook.getSheetAt(0); // Obtener la primera hoja

            // Procesar una gran cantidad de datos
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 0; i < 10000; i++) { // Aumenta este número según tus necesidades
                Row row = sheet.createRow(++lastRowNum);
                Cell cell = row.createCell(0);
                cell.setCellValue("Dato " + (i + 1));


            }
/*
            // Guardar el SXSSFWorkbook como un nuevo archivo XLSX
            try (OutputStream os = new FileOutputStream(outputPath)) {
                workbook.write(os);
            }

            // Dado que estamos usando una versión de streaming, no olvides eliminar los archivos temporales
            workbook.dispose();
            */

            workbook.write(outputStream);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

