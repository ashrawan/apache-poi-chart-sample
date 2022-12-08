package com.projectsample.libapachepoi.playground.temp;

import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

class CreateXSSFColor {

    public static void main(String[] args) throws Exception {

        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fileout = new FileOutputStream("temp-XSSF-color.xlsx")) {

            String rgbS = "FFF000";
            byte[] rgbB = Hex.decodeHex(rgbS); // get byte array from hex string
            XSSFColor color = new XSSFColor(rgbB, null); //IndexedColorMap has no usage until now. So it can be set null.

            XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
            cellStyle.setFillForegroundColor(color);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Sheet sheet = workbook.createSheet();
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("yellow");
            cell.setCellStyle(cellStyle);

            workbook.write(fileout);
        }

    }
}
