package com.projectsample.libapachepoi.chart.generators;

import com.projectsample.libapachepoi.chart.ExcelIndexDTO;
import com.projectsample.libapachepoi.chart.util.ExcelColumn;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class ChartConfigurer {

    // Generic Instances to work on top of XSSFChart
    private XSSFSheet xssfSheet;
    private XSSFChart xssfChart;

    // chart parameters
    ExcelChartProperties.ExcelChartParameters excelChartParameters;
    private ExcelIndexDTO chartIndexDTO;
    private List<ExcelColumn> columns;

    private long seriesStartId = 0;

    // when using low-level api features - use custom.
    // Note: When using custom, all types in combo should use custom, else it seems to break xml
    private boolean useCustomChartCreation = true;

    // Specific instances that are configured along the implementation flow
    private XDDFCategoryAxis categoryAxis;
    private XDDFValueAxis valueAxis;

    public ChartConfigurer() {
    }

    public ChartConfigurer(XSSFSheet xssfSheet, XSSFChart xssfChart, ExcelChartProperties.ExcelChartParameters excelChartParameters, ExcelIndexDTO chartIndexDTO, List<ExcelColumn> columns) {
        this.xssfSheet = xssfSheet;
        this.xssfChart = xssfChart;
        this.excelChartParameters = excelChartParameters;
        this.chartIndexDTO = chartIndexDTO;
        this.columns = columns;
    }

    public XSSFSheet getXssfSheet() {
        return xssfSheet;
    }

    public void setXssfSheet(XSSFSheet xssfSheet) {
        this.xssfSheet = xssfSheet;
    }

    public XSSFChart getXssfChart() {
        return xssfChart;
    }

    public void setXssfChart(XSSFChart xssfChart) {
        this.xssfChart = xssfChart;
    }

    public ExcelChartProperties.ExcelChartParameters getExcelChartParameters() {
        return excelChartParameters;
    }

    public void setExcelChartParameters(ExcelChartProperties.ExcelChartParameters excelChartParameters) {
        this.excelChartParameters = excelChartParameters;
    }

    public ExcelIndexDTO getChartIndexDTO() {
        return chartIndexDTO;
    }

    public void setChartIndexDTO(ExcelIndexDTO chartIndexDTO) {
        this.chartIndexDTO = chartIndexDTO;
    }

    public List<ExcelColumn> getColumns() {
        return columns;
    }

    public void setColumns(List<ExcelColumn> columns) {
        this.columns = columns;
    }

    public long getSeriesStartId() {
        return seriesStartId;
    }

    public void setSeriesStartId(long seriesStartId) {
        this.seriesStartId = seriesStartId;
    }

    public boolean isUseCustomChartCreation() {
        return useCustomChartCreation;
    }

    public void setUseCustomChartCreation(boolean useCustomChartCreation) {
        this.useCustomChartCreation = useCustomChartCreation;
    }

    public XDDFCategoryAxis getCategoryAxis() {
        return categoryAxis;
    }

    public void setCategoryAxis(XDDFCategoryAxis categoryAxis) {
        this.categoryAxis = categoryAxis;
    }

    public XDDFValueAxis getValueAxis() {
        return valueAxis;
    }

    public void setValueAxis(XDDFValueAxis valueAxis) {
        this.valueAxis = valueAxis;
    }
}
