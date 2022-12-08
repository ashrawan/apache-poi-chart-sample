package com.projectsample.libapachepoi.chart.generators.helper;

import com.projectsample.libapachepoi.chart.ExcelIndexDTO;
import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

public class ChartInitializationHelper {

    public static final int CHART_AREA_ROW_SIZE = 16;

    private ChartInitializationHelper() {
    }

    public static XSSFChart initializeChart(SXSSFSheet sXSSFSheet,
                                            ExcelChartProperties.ExcelPosition chartPosition, String chartTitle,
                                            ExcelChartProperties.ExcelPosition legendPosition, int columnSize, ExcelIndexDTO chartIndexDTO) {
        int dataRowStartIndex = chartIndexDTO.getChartRowStartIndex();
        int excelFooterRowEndIndex = chartIndexDTO.getChartRowEndIndex();
        ClientAnchor anchor;
        XSSFDrawing drawing = sXSSFSheet.getDrawingPatriarch();
        switch (chartPosition != null ? chartPosition : ExcelChartProperties.ExcelPosition.TOP_RIGHT) {
            case TOP:
                anchor = drawing
                        .createAnchor(0, 0, 0, 0, 0, dataRowStartIndex - CHART_AREA_ROW_SIZE, 5, dataRowStartIndex - 1);
                break;
            case RIGHT:
                anchor = drawing
                        .createAnchor(0, 0, 0, 0, columnSize + 1, 1,
                                columnSize + 11, 20);
                break;
            case BOTTOM:
                anchor = drawing
                        .createAnchor(0, 0, 0, 0, 0, excelFooterRowEndIndex + 1, 10, 20);
                break;
            case TOP_RIGHT:
            default:
                anchor = drawing
                        .createAnchor(0, 0, 0, 0, 3, 0, 10, dataRowStartIndex - 1);
                break;
        }

        XSSFChart xssfChart = drawing.createChart(anchor);
        xssfChart.setTitleText(chartTitle);
        xssfChart.setTitleOverlay(false);
        XDDFChartLegend legend = xssfChart.getOrAddLegend();

        switch (legendPosition != null ? legendPosition : ExcelChartProperties.ExcelPosition.TOP_RIGHT) {
            case TOP:
                legend.setPosition(LegendPosition.TOP);
                break;
            case LEFT:
                legend.setPosition(LegendPosition.LEFT);
                break;
            case BOTTOM:
                legend.setPosition(LegendPosition.BOTTOM);
                break;
            case RIGHT:
                legend.setPosition(LegendPosition.RIGHT);
                break;
            case TOP_RIGHT:
            default:
                legend.setPosition(LegendPosition.TOP_RIGHT);
                break;
        }

//        legend.setOverlay(false);
//        XDDFManualLayout layout = legend.getOrAddManualLayout();
//        layout.setXMode(LayoutMode.EDGE);
//        layout.setYMode(LayoutMode.EDGE);
//        layout.setX(0.00); //left edge of the chart
//        layout.setY(0.25); //25% of chart's height from top edge of the chart

        return xssfChart;
    }

    // Note: AxisPosition refers to the labels shown on the axis, for category axis labels is usually at the BOTTOM
    public static XDDFCategoryAxis initializeCategoryAxis(XSSFChart xssfChart, String categoryAxisTitle, boolean isCatAxisVisible, AxisPosition categoryAxisPosition) {
        XDDFCategoryAxis categoryAxis;
        switch (categoryAxisPosition != null ? categoryAxisPosition : AxisPosition.BOTTOM) {
            case TOP:
                categoryAxis = xssfChart.createCategoryAxis(AxisPosition.TOP);
                break;
            case LEFT:
                categoryAxis = xssfChart.createCategoryAxis(AxisPosition.LEFT);
                break;
            case RIGHT:
                categoryAxis = xssfChart.createCategoryAxis(AxisPosition.RIGHT);
                break;
            case BOTTOM:
            default:
                categoryAxis = xssfChart.createCategoryAxis(AxisPosition.BOTTOM);
                break;
        }
        if(isCatAxisVisible) {
           categoryAxis.setVisible(true);
           categoryAxis.setTitle(categoryAxisTitle);
        } else {
            categoryAxis.setVisible(false);
        }
        return categoryAxis;
    }

    // Note: AxisPosition.RIGHT does not means that the axis is at the right side of the chart but that the axis labels are right side.
    public static XDDFValueAxis initializeValueAxis(XSSFChart xssfChart,
                                                    String valueAxisTitle,
                                                    AxisPosition valueAxisPosition) {
        XDDFValueAxis valueAxis;
        switch (valueAxisPosition != null ? valueAxisPosition : AxisPosition.BOTTOM) {
            case TOP:
                valueAxis = xssfChart.createValueAxis(AxisPosition.TOP);
                break;
            case LEFT:
                valueAxis = xssfChart.createValueAxis(AxisPosition.LEFT);
                break;
            case RIGHT:
                valueAxis = xssfChart.createValueAxis(AxisPosition.RIGHT);
                break;
            case BOTTOM:
            default:
                valueAxis = xssfChart.createValueAxis(AxisPosition.BOTTOM);
                break;
        }
        valueAxis.setTitle(valueAxisTitle);
        return valueAxis;
    }
}
