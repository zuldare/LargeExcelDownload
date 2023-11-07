package com.jh.largeexceldownload.excel;

import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

@RestController
public class ExcelDownloadController {

    private final Path fileLocation = Paths.get(".");

    @GetMapping("/download/excel")
    public ResponseEntity<Resource> downloadExcel() throws IOException {
        String fileName = "objects" + System.currentTimeMillis() +".xlsx";

        ExcelDownloadService.objectsToExcel("./" + fileName);

        File file = fileLocation.resolve(fileName).toFile();

        if (!file.exists()) {
            throw new FileNotFoundException("El archivo no fue encontrado: " + fileName);
        }

        InputStreamResource resource = new InputStreamResource(new FileInputStream(file));

        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=" + fileName)
                // Asegúrate de setear el Content-Length para que el cliente pueda mostrar el progreso de la descarga
                .contentLength(file.length())
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(resource);
    }
}

