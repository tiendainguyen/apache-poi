package com.example.excel;

import lombok.*;

import java.util.List;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@ToString
class Item {
    private String stt;
    private String maHang;
    private String tenHang;
    private List<SubItem> listSubItem;
    private String donVi;
}
