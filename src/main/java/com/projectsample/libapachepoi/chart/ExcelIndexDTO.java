package com.projectsample.libapachepoi.chart;

import lombok.Data;

@Data
public class ExcelIndexDTO {
    private int dataRowStartIndex;
    private int dataRowEndIndex;

    private int chartRowStartIndex;
    private int chartRowEndIndex;
}
