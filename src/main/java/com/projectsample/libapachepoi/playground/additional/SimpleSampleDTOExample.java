package com.projectsample.libapachepoi.playground.additional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class SimpleSampleDTOExample {

    public static void main(String[] args) throws IOException, IllegalAccessException {
        testSampleDTOExample();
    }

    public static void testSampleDTOExample() throws IOException, IllegalAccessException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Cities & Unique NPCs");
            List<SampleDataDTO> citiesAndNPCsData = new ArrayList<>();

            Map<String, Double> mapWindhelm = new HashMap<>();
            mapWindhelm.put("Nord", 29.0);
            mapWindhelm.put("Altmer", 4.0);
            mapWindhelm.put("Imperial", 6.0);
            mapWindhelm.put("Redguard", 1.0);
            mapWindhelm.put("Breton", 0.0);

            Map<String, Double> mapSolitude = new HashMap<>();
            mapSolitude.put("Nord", 29.0);
            mapSolitude.put("Altmer", 3.0);
            mapSolitude.put("Imperial", 10.0);
            mapSolitude.put("Redguard", 7.0);
            mapSolitude.put("Breton", 5.0);

            Map<String, Double> mapMarkarth = new HashMap<>();
            mapMarkarth.put("Nord", 22.0);
            mapMarkarth.put("Altmer", 2.0);
            mapMarkarth.put("Imperial", 6.0);
            mapMarkarth.put("Redguard", 5.0);
            mapMarkarth.put("Breton", 16.0);

            Map<String, Double> mapRiften = new HashMap<>();
            mapRiften.put("Nord", 35.0);
            mapRiften.put("Altmer", 1.0);
            mapRiften.put("Imperial", 8.0);
            mapRiften.put("Redguard", 3.0);
            mapRiften.put("Breton", 4.0);

            citiesAndNPCsData.add(new SampleDataDTO("Windhelm", Collections.unmodifiableMap(mapWindhelm)));
            citiesAndNPCsData.add(new SampleDataDTO("Solitude", Collections.unmodifiableMap(mapSolitude)));
            citiesAndNPCsData.add(new SampleDataDTO("Markarth", Collections.unmodifiableMap(mapMarkarth)));
            citiesAndNPCsData.add(new SampleDataDTO("Riften", Collections.unmodifiableMap(mapRiften)));

            generateLineChartInExcel(wb, sheet, citiesAndNPCsData);
        }
    }

    public static void generateLineChartInExcel(Workbook workbook, XSSFSheet sheet, List<SampleDataDTO> data) throws IOException, IllegalAccessException {
        int NUM_OF_COLUMNS = 4;

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

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, data.size() + 5, 10, 30);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(sheet.getSheetName());
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);

        // Use a category axis for the bottom axis.
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.LEFT);
        categoryAxis.setTitle("Cities");
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.BOTTOM);
        valueAxis.setTitle("Unique NPCs");
        // BETWEEN category axis crosses the value axis between the strokes and not midpoint the strokes. Else the bars are only half wide visible for first and last category.
        valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

        XDDFDataSource<String> categoryDataSource = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 1, NUM_OF_COLUMNS));

        XDDFChartData chartData = chart.createData(ChartTypes.BAR, categoryAxis, valueAxis);

        int rowIndex = 1;
        for (SampleDataDTO sampleDataDTO: data) {
            XDDFNumericalDataSource<Double> tempValueDataSource = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(rowIndex, rowIndex, 1, NUM_OF_COLUMNS));
            XDDFChartData.Series tempSeries = chartData.addSeries(categoryDataSource, tempValueDataSource);
            tempSeries.setTitle(sampleDataDTO.getLabel(), null);
            rowIndex++;
        }
        chart.plot(chartData);

        // correcting the overlap so bars really are stacked and not side by side
        chart.getCTChart().getPlotArea().getBarChartArray(0).addNewOverlap().setVal((byte) 100);

        // in order to transform a bar chart into a column chart, you just need to change the bar direction
        XDDFBarChartData bar = (XDDFBarChartData) chartData;
        bar.setBarDirection(BarDirection.BAR);
        bar.setBarGrouping(BarGrouping.PERCENT_STACKED);

        // Write the output to a file
        try (FileOutputStream fileOut = new FileOutputStream("ADDITIONAL-Simple-sample-example.xlsx")) {
            workbook.write(fileOut);
        }
    }
}
