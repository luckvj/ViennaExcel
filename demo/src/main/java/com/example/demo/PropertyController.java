package com.example.demo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class PropertyController {

    @GetMapping("/add-property")
    public ResponseEntity<InputStreamResource> addProperty(
            @RequestParam String fullName,
            @RequestParam String address,
            @RequestParam String date,
            @RequestParam String time,
            @RequestParam String price,
            @RequestParam int numBedrooms,
            @RequestParam int numBathrooms) {

        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ExcelWriter.writeToExcel(fullName, address, date, time, price, numBedrooms, numBathrooms, baos);

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fullName + "_PropertyDetails.xlsx");
            headers.add(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
            InputStreamResource resource = new InputStreamResource(bais);

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(resource);

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
}
