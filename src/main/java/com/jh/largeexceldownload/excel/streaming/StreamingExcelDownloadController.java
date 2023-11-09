package com.jh.largeexceldownload.excel.streaming;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;

@RestController
public class StreamingExcelDownloadController {

    @GetMapping("/download/excel")
    public ResponseEntity<Void> downloadExcelFile(HttpServletResponse response) throws IOException {
        // Configura la respuesta HTTP
        String filename = "largefile.xlsx";
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"" + filename + "\"");

        // Genera el Excel y escribe directamente en el OutputStream de la respuesta
        LargeExcelFromTemplate generator = new LargeExcelFromTemplate();
        generator.generate(response.getOutputStream());

        return ResponseEntity.ok().build();
    }

}

