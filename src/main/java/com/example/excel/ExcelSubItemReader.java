package com.example.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelSubItemReader {
    public static List<SubItem> readListSubItem() {
        List<SubItem> subItems = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream("TÍNH LẠI ĐỊNH MỨC NĂM 2023.xlsx");
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                // Skip header or empty rows if necessary
                if (row.getRowNum() <2) continue;
                // Assuming the first three columns contain the data you want
                SubItem subItem = new SubItem();
                subItem.setMaSp(row.getCell(0).toString());
                subItem.setMaKt(row.getCell(1) != null ? row.getCell(1).toString() : null);
                subItem.setTen(row.getCell(2) != null ? row.getCell(2).toString() : null);
                subItem.setDonViTinh(row.getCell(3) != null ? row.getCell(3).toString() : null);
                subItem.setLuongNlThucTeSuDungDeSanXuatMotSanPham((row.getCell(7) != null && !row.getCell(7).toString().isEmpty())?Double.parseDouble(row.getCell(7).toString()) : 0);
                subItems.add(subItem);
            }
            // Now 'items' contains all the objects populated from the Excel file
//            items.forEach(System.out::println);
            System.out.println("size:" + subItems.size());
            System.out.println(String.format("fist item (%s), last item (%s)", subItems.get(0), subItems.get(subItems.size()-2)));
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            return subItems;
        }
    }
}
