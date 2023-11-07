package com.jh.largeexceldownload;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;

public class ImageTest {
    public static void main(String[] args) {
        InputStream is = ImageTest.class.getResourceAsStream("/slack.png");
        if (is == null) {
            System.out.println("El InputStream es null, verifica la ruta del archivo.");
        } else {
            try {
                BufferedImage image = ImageIO.read(is);
                if (image == null) {
                    System.out.println("ImageIO no pudo leer la imagen, verifica que es un archivo PNG válido.");
                } else {
                    System.out.println("La imagen se ha leído correctamente.");
                }
            } catch (IOException e) {
                System.out.println("Ocurrió un error al leer la imagen: " + e.getMessage());
            }
        }
    }
}