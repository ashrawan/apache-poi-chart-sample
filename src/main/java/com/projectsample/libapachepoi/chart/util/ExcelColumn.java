package com.projectsample.libapachepoi.chart.util;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExcelColumn<T> {

    private String colName;

    private int rowStart;
    private int rowEnd;
    private int columnStart;
    private int columnEnd;


}
