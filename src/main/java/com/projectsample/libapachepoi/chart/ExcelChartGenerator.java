package com.projectsample.libapachepoi.chart;

import com.projectsample.libapachepoi.chart.chartmodel.*;
import com.projectsample.libapachepoi.chart.generators.ChartConfigurer;
import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import com.projectsample.libapachepoi.chart.generators.helper.ChartInitializationHelper;
import com.projectsample.libapachepoi.chart.util.ExcelColumn;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

@Slf4j
public class ExcelChartGenerator {


    private SXSSFSheet sheet;
    private ExcelIndexDTO chartIndexDTO;
    private List<ExcelColumn> columns;

    public ExcelChartGenerator(SXSSFSheet sheet) {
        this.sheet = sheet;
    }

    public void drawChart(ExcelChartProperties excelChartProperties) {

        if (excelChartProperties != null) {
            sheet.createDrawingPatriarch();
            XSSFSheet xssfSheet = this.sheet.getDrawingPatriarch().getSheet();
            log.info("Initializing XSSFChart and drawing area for: {}", excelChartProperties.getChartTitle());
            XSSFChart chart = ChartInitializationHelper.initializeChart(this.sheet,
                    excelChartProperties.getChartPosition(),
                    excelChartProperties.getChartTitle(),
                    excelChartProperties.getLegendPosition(),
                    columns.size(), chartIndexDTO);

            XDDFValueAxis primaryValueAxis = null;
            XDDFValueAxis secondaryValueAxis = null;

            // In-Combo chart each series must have different id, so tracking seriesStartId on basis of each-chart dataColumns
            long seriesStartId = 0;
            // once cat axis is initialized, turn off visibility of other cat axis to avoid overlap
            boolean isCatAxisVisible = true;

            for (ExcelChartProperties.ExcelChartParameters chartParams : excelChartProperties
                    .getParams()) {
                log.info("Initializing chartConfigurer and setting up axis for chartType: {}", chartParams.getType());
                ChartConfigurer chartConfigurer = new ChartConfigurer(xssfSheet, chart, chartParams, chartIndexDTO, columns);

                // Initializing categoryAxis for each chart-model
                String categoryAxisTitle = chartParams.getCategoryAxisTitle();
                XDDFCategoryAxis categoryAxis = ChartInitializationHelper.initializeCategoryAxis(chart, categoryAxisTitle, isCatAxisVisible, AxisPosition.BOTTOM);
                isCatAxisVisible = false;

                // Initializing valueAxis - PrimaryAxis (Left Axis) Or SecondaryAxis (Right Axis) as per chart parameters
                XDDFValueAxis valueAxis;
                if (chartConfigurer.getExcelChartParameters().isPlotOnSecondaryAxis()) {
                    if (secondaryValueAxis == null) {
                        String secondaryValueAxisTitle = chartParams.getValueAxisTitle();
                        secondaryValueAxis = ChartInitializationHelper.initializeValueAxis(chart, secondaryValueAxisTitle, AxisPosition.RIGHT);
                        secondaryValueAxis.setCrosses(AxisCrosses.MAX);
                    }
                    if (primaryValueAxis != null && chartConfigurer.getExcelChartParameters().isUseSameMinMaxScaleAsPrimary()) {
                        secondaryValueAxis.setMaximum(primaryValueAxis.getMaximum());
                        secondaryValueAxis.setMinimum(primaryValueAxis.getMinimum());
                    }
                    valueAxis = secondaryValueAxis;
                } else {
                    if (primaryValueAxis == null) {
                        String primaryValueAxisTitle = chartParams.getValueAxisTitle();
                        primaryValueAxis = ChartInitializationHelper.initializeValueAxis(chart, primaryValueAxisTitle, AxisPosition.LEFT);
                    }
                    valueAxis = primaryValueAxis;
                }

                categoryAxis.crossAxis(valueAxis);
                valueAxis.crossAxis(categoryAxis);

                // Setting axis and series sequence for each chart-model
                chartConfigurer.setCategoryAxis(categoryAxis);
                chartConfigurer.setValueAxis(valueAxis);

                chartConfigurer.setSeriesStartId(seriesStartId);
                seriesStartId = seriesStartId + chartParams.getDataRows().size();

                // Instantiating chart-model and drawing it
                ExcelChart excelChart = getExcelChartFromParams(chartParams);
                if (excelChart == null) {
                    log.error("Error parsing type: {} for chartParams request : {}", chartParams.getType(), chartParams);
                    throw new RuntimeException("Error parsing excelChart type: {}" + chartParams.getType());
                }
                log.info("Started drawing-xml with categoryColumns: {} dataColumns: {} for: {} chart",
                        chartParams.getCategoryColumns(), chartParams.getDataRows(), chartParams.getType());
                excelChart.drawChart(chartConfigurer);
            }
        }
    }

    private ExcelChart getExcelChartFromParams(ExcelChartProperties.ExcelChartParameters chartParams) {
        switch (chartParams.getType()) {
            case LINE:
                return new ExcelLineChart();
            case BAR:
                return new ExcelBarChart();
            case PIE:
                return new ExcelPieChart();
            case SCATTER:
                return new ExcelScatterChart();
            default:
                return null;
        }
    }


    public ExcelIndexDTO getChartIndexDTO() {
        if (null == this.chartIndexDTO) {
            this.chartIndexDTO = new ExcelIndexDTO();
        }
        return chartIndexDTO;
    }

    public void setColumns(List<ExcelColumn> columns) {
        this.columns = columns;
    }

//    public void setChartIndexDTO(
//            ExcelChartIndexDTO chartIndexDTO) {
//        this.chartIndexDTO = chartIndexDTO;
//    }


}
