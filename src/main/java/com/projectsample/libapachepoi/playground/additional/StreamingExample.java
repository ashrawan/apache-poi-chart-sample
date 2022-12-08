package com.projectsample.libapachepoi.playground.additional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.*;
import java.util.stream.Collectors;

public class StreamingExample {

    public static void main(String[] args) throws InterruptedException, ExecutionException, IllegalAccessException, IOException {
        testStreamingLargeRecordsData();
    }

    public static void testStreamingLargeRecordsData() throws InterruptedException, ExecutionException, IOException, IllegalAccessException {
        String[] columnCategory = new String[]{"Riften", "Solitude", "Morthal", "Whiterun", "Markarth"};
        ExecutorService taskExecutor = Executors.newFixedThreadPool(4);
        List<Callable<SampleDataDTO>> callables = new ArrayList<>();

        // Initializing Seed Data generator thread
        for (int i = 0; i < 800000; i++) {
            int index = i;
            callables.add(() -> {
                Map<String, Double> tempMap = new HashMap<>();
                for (String cc : columnCategory) {
                    tempMap.put(cc, (double) ThreadLocalRandom.current().nextInt(100));
                }
                SampleDataDTO sampleDataDTO = new SampleDataDTO("data" + index, Collections.unmodifiableMap(tempMap));
                return sampleDataDTO;
            });
        }

        // Executing and retrieving List<SampleDataDTO>
        List<Future<SampleDataDTO>> futures = taskExecutor.invokeAll(callables);
        List<SampleDataDTO> sampleDataDTOS = futures.parallelStream()
                .map(sampleDataDTOFuture -> {
                    try {
                        return sampleDataDTOFuture.get();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    return null;
                }).collect(Collectors.toList());

        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            StreamingExample.generateChartInExcel(wb, "Cities and Area - Last 25", sampleDataDTOS);
        }

    }

    public static void generateChartInExcel(SXSSFWorkbook workbook, String sheetName, List<SampleDataDTO> data) throws IOException, IllegalAccessException {
        System.out.println("Started generateChartInExcel");
        long startTime = System.currentTimeMillis();

        SXSSFSheet sxssfSheet = workbook.createSheet(sheetName);

        int NUM_OF_COLUMNS = 2;

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

        SXSSFDrawing sxssfDrawing = sxssfSheet.createDrawingPatriarch();
        XSSFDrawing drawing = sxssfSheet.getDrawingPatriarch();
        XSSFSheet xssfSheet = drawing.getSheet();

        ClientAnchor anchor = sxssfDrawing.createAnchor(0, 0, 0, 0, NUM_OF_COLUMNS + 2, 0, NUM_OF_COLUMNS + 2 + 15, 20);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(xssfSheet.getSheetName());
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);

        // Use a category axis for the bottom axis.
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        categoryAxis.setTitle("Cities");
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setTitle("Area");
        // BETWEEN category axis crosses the value axis between the strokes and not midpoint the strokes. Else the bars are only half wide visible for first and last category.
        valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

        XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(xssfSheet, new CellRangeAddress(0, 0, 1, NUM_OF_COLUMNS));
        String pointAt = xs.getPointAt(0);
        XDDFChartData chartData = chart.createData(ChartTypes.BAR, categoryAxis, valueAxis);

        // All data sets - data1, data2, data3, data4........ (I am plotting only top 25 data points)
        int dataStartRow = 1;
        for (int i = 800000; i > (800000 - 25); i--) {
            XDDFNumericalDataSource<Double> tempData = XDDFDataSourcesFactory.fromNumericCellRange(xssfSheet, new CellRangeAddress(i, i, 1, NUM_OF_COLUMNS));
            XDDFChartData.Series series1 = chartData.addSeries(xs, tempData);
            series1.setTitle("data" + i, null);
            dataStartRow++;
        }
        chart.plot(chartData);

//        Do this, only if its STACKED bar chart, correcting the overlap so bars really are stacked and not side by side
//        chart.getCTChart().getPlotArea().getBarChartArray(0).addNewOverlap().setVal((byte) 100);

        // in order to transform a bar chart into a column chart, you just need to change the bar direction
        XDDFBarChartData bar = (XDDFBarChartData) chartData;
        bar.setBarDirection(BarDirection.COL);
        bar.setBarGrouping(BarGrouping.CLUSTERED);

        // Write the output to a file
        try (FileOutputStream fileOut = new FileOutputStream("ADDITIONAL-streaming-example.xlsx")) {
            workbook.write(fileOut);
            System.out.println("Completed generating ADDITIONAL-streaming-example.xlsx file");
            long endTime = System.currentTimeMillis();
            System.out.println("That took " + (endTime - startTime) + " milliseconds");
        }
    }
}
