package com.example.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class ExcelWriterLast {
    public static void main(String[] args) {
        // Assuming we already have a method readExcelFile that returns a List<Item> populated with data
        List<Item> items = readExcelFile();
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("sheet1.xlsx");
            int startRow = 11; // Rows in Excel are 0-based; 20 will be the 21st row.
            int startColumn = 1; // Columns in Excel are 0-based; 3 is the fourth column.
            // Adjust for zero-based indexing by decrementing startRow and startColumn
            startRow--;
            startColumn--;
            long number = 0L;
            int rowCounter = startRow;
            for (Item item : items) {
                Row itemRow = sheet.createRow(rowCounter);
                number++;
//                if(number == 798) {
//                    System.out.println("item last: " + item.getListSubItem());
//                }
                itemRow.createCell(startColumn).setCellValue(number);
                itemRow.createCell(startColumn +1).setCellValue(item.getMaHang());
                itemRow.createCell(startColumn + 2).setCellValue(item.getTenHang());
                itemRow.createCell(startColumn + 3).setCellValue(item.getDonVi());
                // Write sub-items
                for (SubItem subItem : item.getListSubItem()) {
//                    if (number == 798) {
//                        System.out.println(subItem);
//                    }
                    rowCounter++;
                    Row subItemRow = sheet.createRow(rowCounter);
                    if (number == 798 && subItem.getMaKt().equals("TUI-001AKV")) {
                        System.out.println(subItem);
                    }
                    subItemRow.createCell(startColumn + 4).setCellValue(subItem.getMaKt());
                    subItemRow.createCell(startColumn + 5).setCellValue(subItem.getTen());
                    subItemRow.createCell(startColumn + 6).setCellValue(subItem.getDonViTinh());
                    subItemRow.createCell(startColumn + 7).setCellValue(subItem.getLuongNlThucTeSuDungDeSanXuatMotSanPham());
                    // ... and so on for the other SubItem properties
                }
                rowCounter++;
            }
            // Writing the workbook to a file
            FileOutputStream out = new FileOutputStream("last.xlsx");
            workbook.write(out);
            out.close();
            System.out.println("Excel file with gaps and items created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private static List<Item> readExcelFile() {
        List<Item> items = ExcelItemReader.readListItem();
        List<SubItem> subItems = ExcelSubItemReader.readListSubItem();
        items = items.stream().map(x -> {
            List<SubItem> sub = new ArrayList<>();
            subItems.stream().forEach( s -> {
                if (s.getMaSp().equals(x.getMaHang())) {
                    if(!s.getMaKt().equals(x.getMaHang())) {
                        sub.add(s);
                    } else {
                        x.setDonVi(s.getDonViTinh());
                    }
                }
            });
            x.setListSubItem(sub);
            if(x.getMaHang().equals("ONDD1-007L1")) {
                System.out.println(x);
            }
            return x;
        }).collect(Collectors.toList());
        // ... Your code to read from the Excel file and populate the items list ...
        return items;
    }
}
