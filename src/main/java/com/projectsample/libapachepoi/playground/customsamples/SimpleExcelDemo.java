package com.projectsample.libapachepoi.playground.customsamples;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class SimpleExcelDemo {

    public static void main(String[] args) {
        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{"ID", "NAME", "Age"});
        data.put("2", new Object[]{1, "James", 25});
        data.put("3", new Object[]{2, "Robert", 32});
        data.put("4", new Object[]{3, "John", 41});
        data.put("5", new Object[]{4, "Michael", 67});
        createDemoExcel(data);
    }

    public static void createDemoExcel(Map<String, Object[]> data) {
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }
        }
        try {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("SAMPLE-simple-excel.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("File general-poi-example.xlsx was successfully created.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
