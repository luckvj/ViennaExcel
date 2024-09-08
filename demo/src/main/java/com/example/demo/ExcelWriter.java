package com.example.demo;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

    public static void writeToExcel(String fullName, String address, String date, String time, String price, int numBedrooms, int numBathrooms, ByteArrayOutputStream baos) throws IOException {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();
        // Create a new sheet
        Sheet sheet = workbook.createSheet("Property Details");

        // Create headers
        Row headerRow = sheet.createRow(0);
        String[] headers = {"Full Name", "Address", "Date", "Time", "Price", "Bedrooms", "Bathrooms"};
        CellStyle headerCellStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setUnderline(Font.U_SINGLE);
        headerCellStyle.setFont(headerFont);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Write data to the new row
        Row dataRow = sheet.createRow(1);
        dataRow.createCell(0).setCellValue(fullName);
        dataRow.createCell(1).setCellValue(address);
        dataRow.createCell(2).setCellValue(date);
        dataRow.createCell(3).setCellValue(time);
        dataRow.createCell(4).setCellValue(price);
        dataRow.createCell(5).setCellValue(numBedrooms);
        dataRow.createCell(6).setCellValue(numBathrooms);

        // Set column widths for each field
        sheet.setColumnWidth(0, 20 * 256); // Column for Full Name
        sheet.setColumnWidth(1, 30 * 256); // Column for Address
        sheet.setColumnWidth(2, 15 * 256); // Column for Date
        sheet.setColumnWidth(3, 15 * 256); // Column for Time
        sheet.setColumnWidth(4, 15 * 256); // Column for Price
        sheet.setColumnWidth(5, 15 * 256); // Column for Bedrooms
        sheet.setColumnWidth(6, 15 * 256); // Column for Bathrooms        

        // Write the workbook to the ByteArrayOutputStream
        workbook.write(baos);
        workbook.close();
    }
}
