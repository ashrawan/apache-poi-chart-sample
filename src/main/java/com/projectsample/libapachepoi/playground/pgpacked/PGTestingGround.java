package com.projectsample.libapachepoi.playground.pgpacked;

import com.projectsample.libapachepoi.playground.additional.SampleDataDTO;
import com.projectsample.libapachepoi.playground.temp.ChartIndexInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class PGTestingGround {

    public static void main(String[] args) throws IOException, InterruptedException {

        try (SXSSFWorkbook wb = new SXSSFWorkbook(null, 100, true)) {
            SXSSFSheet sxssfSheet = wb.createSheet("pg-cities-npcs");
            PGTestingGround PGTestingGround = new PGTestingGround();
            List<SampleDataDTO> sampleDataDTOS = PGTestingGround.sampleDataSeed();
            int NUM_OF_COLUMNS = PGTestingGround.writingPGData(sxssfSheet, sampleDataDTOS);
            PGTestingGround.creatingPGChart(wb, sxssfSheet, NUM_OF_COLUMNS);
        }
    }

    private List<SampleDataDTO> sampleDataSeed() {
        List<SampleDataDTO> citiesAndNPCsData = new ArrayList<>();

        Map<String, Double> mapWindhelm = new HashMap<>();
        mapWindhelm.put("Nord", 30.0);
        mapWindhelm.put("Altmer", 4.0);
        mapWindhelm.put("Imperial", 6.0);
        mapWindhelm.put("Redguard", 1.0);
        mapWindhelm.put("Breton", 0.0);

        Map<String, Double> mapSolitude = new HashMap<>();
        mapSolitude.put("Nord", 24.0);
        mapSolitude.put("Altmer", 3.0);
        mapSolitude.put("Imperial", 10.0);
        mapSolitude.put("Redguard", 7.0);
        mapSolitude.put("Breton", 5.0);

        citiesAndNPCsData.add(new SampleDataDTO("Windhelm", Collections.unmodifiableMap(mapWindhelm)));
        citiesAndNPCsData.add(new SampleDataDTO("Solitude", Collections.unmodifiableMap(mapSolitude)));
        return citiesAndNPCsData;
    }

    private int writingPGData(SXSSFSheet sxssfSheet, List<SampleDataDTO> data) {
        int NUM_OF_COLUMNS = 0;

        Row row;
        Cell cell;

        for (int rowIndex = 0; rowIndex <= data.size(); rowIndex++) {
            row = sxssfSheet.createRow(rowIndex);

            // Populating header in 1st column
            if (rowIndex == 0) {
                Map<String, ?> valuesMap = data.get(rowIndex).getValuesMap();
                NUM_OF_COLUMNS = NUM_OF_COLUMNS > valuesMap.size() ? NUM_OF_COLUMNS : valuesMap.size();
                int colIndex = 1;
                for (Map.Entry<String, ?> entry : valuesMap.entrySet()) {
                    cell = row.createCell(colIndex);
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
                    cell = row.createCell(colIndex);
                    if (rowIndex == 0) {
                        cell.setCellValue(entry.getKey());
                    } else {
                        cell.setCellValue((Double) entry.getValue());
                    }
                    colIndex++;
                }
            }

        }

        return NUM_OF_COLUMNS;
    }

    public void creatingPGChart(SXSSFWorkbook workbook, SXSSFSheet sxssfSheet, int NUM_OF_COLUMNS) throws IOException {
        ChartTypes chartType = ChartTypes.BAR; // (For now plotting BAR Chart as sample)

        // Initializing Drawing Patriarch and setting chart Title and Legend Position
        sxssfSheet.createDrawingPatriarch();
        XSSFSheet xssfSheet = sxssfSheet.getDrawingPatriarch().getSheet();
        XSSFChart xssfChart = PGInitializerUtil.initializeChartDrawingPatriarch
                (xssfSheet, 0, 0, 0, 0, NUM_OF_COLUMNS + 2, 0, NUM_OF_COLUMNS + 2 + 15, 20);
        PGInitializerUtil.initializeChartTitleAndLegend(xssfChart, "Title - PG Cities NPCs Count", LegendPosition.RIGHT);

        // Category and Value Axis
        XDDFCategoryAxis categoryAxis = PGInitializerUtil.initializeCategoryAxis(xssfChart, "Category", AxisPosition.LEFT);
        XDDFValueAxis valueAxis = PGInitializerUtil.initializeValueAxis(xssfChart, "Value in Number", AxisPosition.BOTTOM);

        // To dynamically evaluate row & column for data series in either - horizontal/row order or vertical/column order
        boolean isSeriesLayoutHorizontal = true;
        int maxDataSizeToPlotIntoChart = 2;
        int DATA_START_ROW = isSeriesLayoutHorizontal ? 1 : 1;
        int DATA_END_ROW = isSeriesLayoutHorizontal ? maxDataSizeToPlotIntoChart : 5; // (since, for now we have only populated 5 columns mock data)
        int DATA_START_COLUMN = isSeriesLayoutHorizontal ? 1 : 1;
        int DATA_END_COLUMN = isSeriesLayoutHorizontal ? maxDataSizeToPlotIntoChart : 5; // (since, for now we have only populated 5 columns mock data)

        int CATEGORY_START_ROW = isSeriesLayoutHorizontal ? 0 : 1;
        int CATEGORY_END_ROW = isSeriesLayoutHorizontal ? 0 : 20; // for vertical series, taking 20 rows as category,
        int CATEGORY_START_COLUMN = isSeriesLayoutHorizontal ? 1 : 0;
        int CATEGORY_END_COLUMN = isSeriesLayoutHorizontal ? NUM_OF_COLUMNS : 0;
        ChartIndexInfo chartIndexInfo = new ChartIndexInfo();
        chartIndexInfo.setDataStartRow(DATA_START_ROW);
        chartIndexInfo.setDataEndRow(DATA_END_ROW);
        chartIndexInfo.setDataStartColumn(DATA_START_COLUMN);
        chartIndexInfo.setDataEndColumn(DATA_END_COLUMN);
        chartIndexInfo.setCategoryStartRow(CATEGORY_START_ROW);
        chartIndexInfo.setCategoryEndRow(CATEGORY_END_ROW);
        chartIndexInfo.setCategoryStartColumn(CATEGORY_START_COLUMN);
        chartIndexInfo.setCategoryEndColumn(CATEGORY_END_COLUMN);

        // Specifying CategoryDataSource CellRange (uses single CellRange set) and SeriesValues CellRange (uses CellRange for evaluating multiple data series)
        CellRangeAddress catCellRangeAddress = new CellRangeAddress(chartIndexInfo.getCategoryStartRow(), chartIndexInfo.getCategoryEndRow(), chartIndexInfo.getCategoryStartColumn(), chartIndexInfo.getCategoryEndColumn());
        XDDFDataSource<?> xddfCategoryDataSource = PGInitializerUtil.setCategoryDataSourceUsingCellRange(xssfSheet, false, catCellRangeAddress);
        // Creates chartData
        XDDFChartData chartData = PGInitializerUtil.createXDDFChartData(xssfChart, chartType, categoryAxis, valueAxis);

        XDDFBarChartData xddfBarChartData = (XDDFBarChartData) chartData;
        xddfBarChartData.setBarGrouping(BarGrouping.CLUSTERED);
        xddfBarChartData.setBarDirection(BarDirection.COL
        );
        PGInitializerUtil.addDataAsSeriesUsingCellRange(chartData, xssfSheet, chartIndexInfo, xddfCategoryDataSource, isSeriesLayoutHorizontal);

        xssfChart.plot(chartData);

        // writing the output to file, its only for testing purpose
        try (FileOutputStream fileOut = new FileOutputStream("PG-PlayGround-Sample.xlsx")) {
            workbook.write(fileOut);
            System.out.println("Completed generating PlayGround-Sample.xlsx file");
        }

    }


}
