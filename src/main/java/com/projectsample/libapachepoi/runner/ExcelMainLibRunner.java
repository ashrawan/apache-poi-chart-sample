package com.projectsample.libapachepoi.runner;

import com.projectsample.libapachepoi.chart.ExcelChartGenerator;
import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import com.projectsample.libapachepoi.chart.util.ChartUtils;
import com.projectsample.libapachepoi.chart.util.ExcelColumn;
import com.projectsample.libapachepoi.playground.additional.SampleDataDTO;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;

import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.UUID;

@Slf4j
@Component
public class ExcelMainLibRunner {

    public void exportAsExcel(List<ExcelColumn> columns, List<SampleDataDTO> sampleDataDTOS, ExcelChartProperties excelChartProperties, String sheetName) {
        // Create Workbook
        SXSSFWorkbook workbook = new SXSSFWorkbook(null, SXSSFWorkbook.DEFAULT_WINDOW_SIZE, true);
        ChartUtils.setThemeFromTemplateFile(workbook);
        SXSSFSheet sheet = workbook.createSheet(sheetName);

        // Write tabular data
        generateExcelTableData(sheet, sampleDataDTOS);

        // Configure all columns and data start & end index
        ExcelChartGenerator chartGenerator = new ExcelChartGenerator(sheet);
        chartGenerator.setColumns(columns);

        // set chart population index
        chartGenerator.getChartIndexDTO().setChartRowStartIndex(sheet.getLastRowNum());
        chartGenerator.getChartIndexDTO().setChartRowEndIndex(sheet.getLastRowNum()+1);

        // draw Chart
        chartGenerator.drawChart(excelChartProperties);

        writeExcelFile(workbook, null);
    }


    private void writeExcelFile(SXSSFWorkbook workbook, String filename) {
        if (!StringUtils.hasText(filename)) {
            filename = UUID.randomUUID().toString() + ".xlsx";
        }
        File excelFile = new File(filename);

        try (OutputStream enclosedStream = new BufferedOutputStream(new FileOutputStream(excelFile));) {
            log.debug("Started writing data to excel file {}", excelFile.getAbsolutePath());
            workbook.write(enclosedStream);
            workbook.dispose();
            log.debug("Completed writing to excel file {} ", excelFile.getAbsolutePath());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void generateExcelTableData(SXSSFSheet sheet, List<SampleDataDTO> data) {
        int NUM_OF_COLUMNS = data.size();

        Row row;
        Cell cell;

        for (int rowIndex = 0; rowIndex <= data.size(); rowIndex++) {
            row = sheet.createRow((short) rowIndex);

            // Populating header in 1st column
            if (rowIndex == 0) {
                Map<String, ?> valuesMap = data.get(rowIndex).getValuesMap();
                NUM_OF_COLUMNS = NUM_OF_COLUMNS > valuesMap.size() ? NUM_OF_COLUMNS : valuesMap.size();
                int colIndex = 1;
                for (Map.Entry<String, ?> entry : valuesMap.entrySet()) {
                    cell = row.createCell((short) colIndex);
                    cell.setCellValue(entry.getKey());
                    colIndex++;
                }
            } else {
                Cell cell1 = row.createCell(0);
                cell1.setCellValue(data.get(rowIndex - 1).getLabel());

                // Populating data values
                Map<String, ?> valuesMap = data.get(rowIndex - 1).getValuesMap();
                int colIndex = 1;
                for (Map.Entry<String, ?> entry : valuesMap.entrySet()) {
                    cell = row.createCell((short) colIndex);
                    if (rowIndex == 0) {
                        cell.setCellValue(entry.getKey());
                    } else {
                        cell.setCellValue((Double) entry.getValue());
                    }
                    colIndex++;
                }
            }

        }
    }

}
