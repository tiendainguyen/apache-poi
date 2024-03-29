package com.example.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileOutputStream;
import java.io.IOException;

@SpringBootApplication
public class ExcelWriter1 {

    public static void main(String[] args) {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Custom Sheet");

            // Starting indices accounting for margins
            int startRow = 11; // Excel rows are 0-based; 20 will actually be the 21st row.
            int startColumn = 0; // Excel columns are 0-based; 3 is the fourth column.
            int lastNumber = 1600;

            // Loop to fill the sequence numbers
            for (int i = 1; i <= lastNumber; i++) {
                int rowNumber = startRow + ((i - 1) * 11); // Increase by 11 since we have 11 lines per number
                sheet.createRow(rowNumber).createCell(startColumn).setCellValue(i);
                // Adding 10 empty rows
                for (int j = 1; j <= 10; j++) {
                    sheet.createRow(rowNumber + j);
                }
            }

            // Writing the workbook to a file
            FileOutputStream out = new FileOutputStream("abc.xlsx");
            workbook.write(out);
            out.close();
            System.out.println("Excel file created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
