package com.projectsample.libapachepoi.playground.pgpacked;

import com.projectsample.libapachepoi.playground.temp.ChartIndexInfo;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class PGInitializerUtil {

    public static XDDFChartData createXDDFChartData(XSSFChart chart, ChartTypes chartType, XDDFChartAxis categoryAxis, XDDFValueAxis valueAxis) {
        XDDFChartData chartData;
        if (ChartTypes.PIE.equals(chartType)) {
            chartData = chart.createData(chartType, null, null);
        } else {
            chartData = chart.createData(chartType, categoryAxis, valueAxis);
        }
        return chartData;
    }

    public static XSSFChart initializeChartDrawingPatriarch(XSSFSheet xssfSheet, int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2) {
        XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
        XSSFChart chart = drawing.createChart(anchor);
        return chart;
    }

    public static XDDFChartLegend initializeChartTitleAndLegend(XSSFChart chart, String chartTitle, LegendPosition legendPosition) {
        // Chart title
        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);
        // Legend
        XDDFChartLegend xddfChartLegend = chart.getOrAddLegend();
        xddfChartLegend.setPosition(legendPosition);
        return xddfChartLegend;
    }

    public static XDDFCategoryAxis initializeCategoryAxis(XSSFChart chart, String categoryAxisTitle, AxisPosition categoryAxisPosition) {
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(categoryAxisPosition);
        categoryAxis.setTitle(categoryAxisTitle);
        return categoryAxis;
    }

    public static XDDFValueAxis initializeValueAxis(XSSFChart chart, String valueAxisTitle, AxisPosition valueAxisPosition) {
        XDDFValueAxis valueAxis = chart.createValueAxis(valueAxisPosition);
        valueAxis.setTitle(valueAxisTitle);
        return valueAxis;
    }

    public static XDDFDataSource<?> setCategoryDataSourceUsingCellRange(XSSFSheet xssfSheet, boolean isNumeric, CellRangeAddress categoryCellAddresses) {
        XDDFDataSource<?> categoryDataSource;
        if (isNumeric) {
            categoryDataSource = XDDFDataSourcesFactory.fromNumericCellRange(xssfSheet, categoryCellAddresses);
        } else {
            categoryDataSource = XDDFDataSourcesFactory.fromStringCellRange(xssfSheet, categoryCellAddresses);
        }
        return categoryDataSource;

    }


    /**
     * This method adds data as series to the chart
     * <p>
     * =======================================
     * HORIZONTAL SERIES (this example has 5 categories)
     * <p>
     * (Category: C1, C2, C3, C4, C5 AND Data-Series: d0, d1, d2, d3)
     * <p>
     *      C1	C2	C3	C4	C5	 CellRange: (firstRow, lastRow, firstCol, lastCol)
     * d0	19	28	32	7	53   	series0	(1, 1, 1, 5)
     * d1	20	54	60	88	13		series1	(2, 2, 1, 5)
     * d2	18	50	21	42	85		series2	(3, 3, 1, 5)
     * d3	61	51	39	78	13		series3	(4, 4, 1, 5)
     * ....
     * ....
     * i.e  (i, i, 1, totalCategories/dataColumns.size/5)
     * (i, i, columnStart, totalCategories/dataColumns.size)
     * ====================================================
     * <p>
     * VERTICAL SERIES (this example has 4 categories)
     * <p>
     *      d1	d2	d3	d4	d5
     * C1	19	28	32	7	53
     * C2	20	54	60	88	13
     * C3	18	50	21	42	85
     * C4	61	51	39	78	13
     * <p>
     * (firstRow, lastRow, firstCol, lastCol)
     * d1: series0	(1, 4, 1, 1)
     * d2: series1 (1, 4, 2, 2)
     * d3: series3 (1, 4, 3, 3)
     * <p>
     * i.e (1, totalCategories/dataColumns.size/4, i, i)
     * (dataRowStart, totalCategories/dataColumns.size, i, i)
     * ========================================================
     *
     */
    public static void addDataAsSeriesUsingCellRange(XDDFChartData chartData, XSSFSheet xssfSheet, ChartIndexInfo chartIndexInfo, XDDFDataSource<?> categoryDataSourceValues, boolean isSeriesLayoutHorizontal) {
        int initialFirstRow = 0, initialLastRow = 0, initialFirstCol = 0, initialLastCol = 0;
        int totalNumberOfDataSeriesToPlot = 0;
        int categoriesSize = 0; //totalCategories/dataColumns.size

        // For horizontal-series: no increment required in column AND For vertical-series: no increment required in row
        if (isSeriesLayoutHorizontal) {
            initialFirstRow = chartIndexInfo.getDataStartRow();
            totalNumberOfDataSeriesToPlot = (chartIndexInfo.getDataEndRow() - chartIndexInfo.getDataStartRow()) + 1;
            categoriesSize = (chartIndexInfo.getCategoryEndColumn() - chartIndexInfo.getCategoryStartColumn()) + 1; // if data is in horizontal (row-by-row), category must be in column
        } else {
            initialFirstCol = chartIndexInfo.getDataStartColumn();
            totalNumberOfDataSeriesToPlot = (chartIndexInfo.getDataEndColumn() - chartIndexInfo.getDataStartColumn()) + 1;
            categoriesSize = (chartIndexInfo.getCategoryEndRow() - chartIndexInfo.getCategoryStartRow()) + 1; // if data is in vertical (col-by-col), category must be in row
        }

        for (int i = 0; i < totalNumberOfDataSeriesToPlot; i++) {
            CellRangeAddress cellAddresses = null;
            if (isSeriesLayoutHorizontal) {
                cellAddresses = new CellRangeAddress(initialFirstRow, initialFirstRow, chartIndexInfo.getDataStartColumn(), categoriesSize);
                initialFirstRow++;
            } else {
                cellAddresses = new CellRangeAddress(chartIndexInfo.getDataStartRow(), categoriesSize, initialFirstCol, initialFirstCol);
                initialFirstCol++;
            }
            XDDFNumericalDataSource<Double> tempData = XDDFDataSourcesFactory.fromNumericCellRange(xssfSheet, cellAddresses);

            XDDFChartData.Series series1 = chartData.addSeries(categoryDataSourceValues, tempData);
            // If data was in horizontal ROW then series is ROWCat{index}
            String seriesTitle = isSeriesLayoutHorizontal ? "SeriesB" + i : "SeriesC" + i;
            series1.setTitle(seriesTitle, null);
        }
    }

}
