package com.example.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelItemReader {
    public static List<Item> readListItem() {
        List<Item> items = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream("BAO CAO THANH PHAM 2023 ( BAN TIEP NHAN).xlsx");
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                // Skip header or empty rows if necessary
                if (row.getRowNum() <3) continue;
                // Assuming the first three columns contain the data you want
                Item item = new Item();
                item.setStt(row.getCell(0).toString());
                item.setMaHang(row.getCell(1).toString());
                item.setTenHang(row.getCell(2).toString());
                items.add(item);
                if (item.getMaHang().equals("ONDD1-007")) {
                    System.out.println(item);
                }
            }
            // Now 'items' contains all the objects populated from the Excel file
//            items.forEach(System.out::println);
            System.out.println("size:" + items.size());
            System.out.println(String.format("fist item (%s), last item (%s)", items.get(0), items.get(items.size()-1)));
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            return items;
        }
    }
}
