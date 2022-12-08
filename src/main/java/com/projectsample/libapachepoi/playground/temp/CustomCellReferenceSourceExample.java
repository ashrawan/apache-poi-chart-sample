package com.projectsample.libapachepoi.playground.temp;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

// TODO: Testing Some Customizations
public class CustomCellReferenceSourceExample {

    public static void main(String[] args) throws IOException {
        testCustomCellReferenceExample();
    }

    private static final Random RNG = new Random();

    public static void testCustomCellReferenceExample() throws IOException {

        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            SXSSFSheet sxssfSheet = wb.createSheet("Sheet1");

            SXSSFRow row = sxssfSheet.createRow(0);
            row.createCell(0).setCellValue("Forest Density");
            row.createCell(1).setCellValue("Roadmap Construct");
            row.createCell(2).setCellValue("River Area");
            row.createCell(3).setCellValue("Heat Map");

            String[] monthArr = {"Jan-Mar", "Apr-Jun", "Jul-Sep", "Oct-Dec"};

            SXSSFCell cell;
            int iQuarter = 0;
            int iYear = 2016;
            for (int r = 1; r <= 7; r++) {
                row = sxssfSheet.createRow(r);

                cell = row.createCell(0);
                String tValue = iYear + "-" + monthArr[iQuarter];
                cell.setCellValue(tValue);

                cell = row.createCell(1);
                int densityValue = RNG.nextInt(500) + 100;
                cell.setCellValue((int) densityValue);
                cell = row.createCell(2);

                int densityDifference = RNG.nextInt(50) * 1;
                cell.setCellValue((int) densityDifference);

                cell = row.createCell(3);
                double calulatedPercentage = (densityDifference / 100.00) * densityValue;
                cell.setCellValue((int) densityValue + calulatedPercentage);

                if (iQuarter >= 3) {
                    iQuarter = 0;
                    iYear++;
                } else {
                    iQuarter++;
                }
            }

            row = sxssfSheet.createRow(8);
            row = sxssfSheet.createRow(9);
            row = sxssfSheet.createRow(10);
            cell = row.createCell(0);
            cell.setCellValue(255.0);


            sxssfSheet.createDrawingPatriarch();
            XSSFDrawing drawing = sxssfSheet.getDrawingPatriarch();
            ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 7 + 5, 10, 30);
            XSSFSheet sheet = drawing.getSheet();


            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText(sxssfSheet.getSheetName());
            chart.setTitleOverlay(false);
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.BOTTOM);

            // Use a category axis for the bottom axis.
            XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            categoryAxis.setTitle("Context");
            XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.BOTTOM);
            valueAxis.setTitle("Value");
            // BETWEEN category axis crosses the value axis between the strokes and not midpoint the strokes. Else the bars are only half wide visible for first and last category.
            valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

            // ########################## BAR ##############################
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 7, 0, 0));
            XDDFChartData barChartData = chart.createData(ChartTypes.BAR, categoryAxis, valueAxis);

            XDDFNumericalDataSource<Double> frDensity = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 8, 1, 1));
            XDDFChartData.Series seriesBar1 = barChartData.addSeries(xs, frDensity);
            seriesBar1.setTitle("Forest Density", null);
            XDDFNumericalDataSource<Double> roadDensity = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 8, 2, 2));
            XDDFChartData.Series seriesBar2 = barChartData.addSeries(xs, roadDensity);
            seriesBar2.setTitle("Forest Density", null);

            // Bar grouping and direction
            XDDFBarChartData bar = (XDDFBarChartData) barChartData;
            bar.setBarDirection(BarDirection.COL);
            bar.setBarGrouping(BarGrouping.CLUSTERED);
            chart.plot(barChartData);
            // #################################################################


            // ##################################### Line ##############################
//            XDDFDataSource<String> xsLine = XDDFDataSourcesFactory.fromStringCellRange(sxssfSheet, new CellRangeAddress(1, 7, 0, 0));
//            Double[] heatmaps = {255.0, 255.0, 255.0, 255.0, 255.0, 255.0, 255.0};
//            XDDFNumericalDataSource<Double> heatMapValueData = XDDFDataSourcesFactory.fromArray(heatmaps);

            List<CellReference> cellReferenceList = new ArrayList<>();
            CellReference cf = new CellReference(sxssfSheet.getSheetName(), 1, 0, true,true); // 11,0
            for(int i = 0; i <=7; i++) {
                cellReferenceList.add(cf);
            }
            CustomCellReferenceDataSource customCellReferenceDataSource = new CustomCellReferenceDataSource(sxssfSheet, cellReferenceList);

            XDDFChartData lineChartData = chart.createData(ChartTypes.LINE, categoryAxis, valueAxis);
            XDDFChartData.Series seriesLine1 = lineChartData.addSeries(xs, customCellReferenceDataSource);
            seriesLine1.setTitle("HeatMap", null);
            chart.plot(lineChartData);
            // #################################################################

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream("temp-custom-cellref-example.xlsx")) {
                wb.write(fileOut);
            }

        }

    }


}
