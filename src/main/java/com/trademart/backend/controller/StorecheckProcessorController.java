package com.trademart.backend.controller;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;
import java.util.UUID;

@RestController
@RequestMapping("/api/storecheck")
@CrossOrigin(origins = "*")
public class StorecheckProcessorController {

    @PostMapping(value = "/process", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<Resource> processStorecheck(
            @RequestParam("file") MultipartFile file,
            @RequestParam("config") String configJson) {

        try {
            String tempId = UUID.randomUUID().toString();
            Path tempDirPath = Files.createTempDirectory("temp_sc_" + tempId);
            File tempDir = tempDirPath.toFile();

            // Guardar excel base usando la ruta absoluta para que transferTo no falle
            File baseFile = new File(tempDir, "base.xlsx");
            file.transferTo(baseFile.getAbsoluteFile());

            // Generar archivo JSON de configuración
            ObjectMapper mapper = new ObjectMapper();
            Map<String, Object> configObj = mapper.readValue(configJson, Map.class);
            
            // Actualizar rutas en el JSON
            configObj.put("base_file", baseFile.getAbsolutePath());
            File outputFile = new File(tempDir, "output.xlsx");
            configObj.put("output_file", outputFile.getAbsolutePath());

            File configFile = new File(tempDir, "config.json");
            mapper.writeValue(configFile, configObj);

            // Ejecutar Python
            ProcessBuilder pb = new ProcessBuilder("python", "headless_storecheck.py", configFile.getAbsolutePath());
            pb.redirectErrorStream(true);
            Process process = pb.start();

            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            StringBuilder output = new StringBuilder();
            while ((line = reader.readLine()) != null) {
                output.append(line).append("\n");
            }
            int exitCode = process.waitFor();

            if (exitCode != 0 || !outputFile.exists()) {
                throw new RuntimeException("Error en Python: " + output.toString());
            }

            Resource resource = new FileSystemResource(outputFile);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=storecheck_generado.xlsx")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(resource);

        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.internalServerError().build();
        }
    }
}
