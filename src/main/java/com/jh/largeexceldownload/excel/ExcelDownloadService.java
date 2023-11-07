package com.jh.largeexceldownload.excel;

import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class ExcelDownloadService {


    public static final int NUMBER_OF_COLUMNS = 300;
    public static final int NUMBER_OF_ROWS = 40000;

    public static void objectsToExcel(String filePath) throws IOException {
        // Mantener 100 filas en memoria, las filas excedentes serán escritas en el disco
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(200)) {
            SXSSFSheet sheet = workbook.createSheet("Objects");

            // Añadir una imagen en la fila 0, celda 1 (B1)
            //InputStream imageInputStream = ExcelDownloadService.class.getResourceAsStream("/slack.png");
            InputStream imageInputStream = ExcelDownloadService.class.getResourceAsStream("/ironman.png");
            byte[] bytes = IOUtils.toByteArray(imageInputStream);

            int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
            imageInputStream.close();


            Drawing<?> drawing = sheet.createDrawingPatriarch();

            // Configurar el ancla para la imagen
            XSSFClientAnchor anchor = new XSSFClientAnchor();

            anchor.setCol1(1);
            anchor.setRow1(0);
            anchor.setCol2(2);
            anchor.setRow2(1);

            anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);

            Picture pict = drawing.createPicture(anchor, pictureIdx);
           // pict.resize();

            // Configurar la altura de la fila para la imagen
            Row pictureRow = sheet.createRow(0);
            pictureRow.setHeightInPoints(90);


            // *****************************
            // CREAR UN TITULO LARGO
            // *****************************

            // Crear un estilo de fuente para el texto
            Font font = workbook.createFont();
            font.setFontHeightInPoints((short) 16);
            font.setBold(true);

            CellStyle styleTitleLarge = workbook.createCellStyle();
            styleTitleLarge.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            styleTitleLarge.setAlignment(HorizontalAlignment.CENTER);
            styleTitleLarge.setLocked(true);
            styleTitleLarge.setFont(font);

            Cell cellTitleLarge = pictureRow.createCell(3);
            cellTitleLarge.setCellValue("Título Largo puesto para probar cositas que las veas");
            cellTitleLarge.setCellStyle(styleTitleLarge);
            // *****************************






            // Crear un estilo para las cabeceras
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.BLACK.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setLocked(true);

            // Crear una fuente para las cabeceras
            Font headerFont = workbook.createFont();
            headerFont.setFontName("Calibri");
            headerFont.setFontHeightInPoints((short) 12);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);


            // Crear un estilo para las celdas
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
            cellStyle.setLocked(false);

            // Crear una fuente para las celdas
            Font cellFont = workbook.createFont();
            cellFont.setFontName("Calibri");
            cellFont.setFontHeightInPoints((short) 11);
            cellFont.setColor(IndexedColors.BLACK.getIndex());
            cellFont.setBold(false);
            cellStyle.setFont(cellFont);


            // Crear el header
            Row headerRow = sheet.createRow(2);
            Cell headerCell;
             for(int i = 0; i<= NUMBER_OF_COLUMNS; i++){
                 headerCell = headerRow.createCell(i);
                 headerCell.setCellValue("Nombre del Atributo " + i);
                 headerCell.setCellStyle(headerStyle);
             }



            // Llenar las filas con los objetos
            int rowIndex = 1;
            Row row;
            Cell cell;
            for (rowIndex=3; rowIndex< NUMBER_OF_ROWS; rowIndex++) {
                row = sheet.createRow(rowIndex++);

                for (int i = 0; i <= NUMBER_OF_COLUMNS; i++) {
                    cell = row.createCell(i);
                    cell.setCellValue("Valor " + i);
                    cell.setCellStyle(cellStyle);
                }
            }

            // Ajusta el tamaño de las columnas
            // PROBAR RENDIMIENTO CON Y SIN ESTO
    //         for(int i = 0; i <=NUMBER_OF_COLUMNS; i++) {
    //             sheet.trackAllColumnsForAutoSizing();
    //            sheet.autoSizeColumn(i);
    //        }

            sheet.protectSheet("password"); // Reemplaza "password" con la contraseña que desees


            // Escribe el archivo en el disco
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
            // Limpia los archivos temporales
            workbook.dispose();
        }
    }
}
