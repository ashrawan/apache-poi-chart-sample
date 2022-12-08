package com.projectsample.libapachepoi.chart.chartmodel;

import com.projectsample.libapachepoi.chart.generators.ChartConfigurer;
import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import com.projectsample.libapachepoi.chart.generators.helper.ChartCreationHelper;
import com.projectsample.libapachepoi.chart.generators.helper.ChartCustomizationHelper;
import org.apache.poi.xddf.usermodel.chart.Grouping;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;


public class ExcelLineChart implements ExcelChart {

    @Override
    public void drawChart(ChartConfigurer chartConfigurer) {
        XDDFChartData chartData = ChartCreationHelper.createChartData(chartConfigurer);

        if (chartConfigurer.isUseCustomChartCreation()) {
            customLineCategoryAndSeriesDataSource(chartConfigurer);
        }
        ChartCreationHelper.initAndPlotChart(chartConfigurer, chartData);

        setLineChartStyleGrouping(chartData, Grouping.STANDARD);
        setLineChartSeriesStylesAndLabels(chartConfigurer);

    }

    private CTLineChart getCTLineChart(ChartConfigurer chartConfigurer) {
        // Get last created bar chart element
        CTPlotArea plotArea = chartConfigurer.getXssfChart().getCTChart().getPlotArea();
        int totalLineChartElements = plotArea.sizeOfLineChartArray();
        return plotArea.getLineChartArray(totalLineChartElements - 1);
    }

    private void customLineCategoryAndSeriesDataSource(ChartConfigurer chartConfigurer) {

        CTLineChart ctLineChart = getCTLineChart(chartConfigurer);
        CTBoolean ctBooleanVaryColor = ctLineChart.isSetVaryColors() ? ctLineChart.getVaryColors() : ctLineChart.addNewVaryColors();
        ctBooleanVaryColor.setVal(true);

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();

        for (int i = 0; i < excelChartParameters.getDataRows().size(); i++) {
            CTLineSer ctLineSer = ctLineChart.addNewSer();
            ctLineSer.addNewIdx().setVal(chartConfigurer.getSeriesStartId() + i);

            // ======== DEFINE CATEGORY Data source ==========
            CTAxDataSource cttAxDataSource = ctLineSer.addNewCat();
            ChartCreationHelper.createCategoryCustomDataSource(chartConfigurer,
                    excelChartParameters.getCategoryColumns(),
                    cttAxDataSource);

            // ============ DEFINE EACH SERIES data source ===============
            CTSerTx ctSerTx = ctLineSer.addNewTx();
            CTNumDataSource ctNumDataSource = ctLineSer.addNewVal();
            ChartCreationHelper.createEachSeriesCustomDataSource(chartConfigurer,
                    excelChartParameters.getDataRows().get(i),
                    ctSerTx, ctNumDataSource);

        }

    }

    private void setLineChartStyleGrouping(XDDFChartData xddfChartData, Grouping lineGrouping) {
        XDDFLineChartData xddfLineChartData = (XDDFLineChartData) xddfChartData;
        xddfLineChartData.setGrouping(lineGrouping);
    }

    private void setLineChartSeriesStylesAndLabels(ChartConfigurer chartConfigurer) {
        CTLineChart ctLineChart = getCTLineChart(chartConfigurer);
        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();

        for (int seriesIndex = 0; seriesIndex < ctLineChart.getSerList().size(); seriesIndex++) {

            CTLineSer ctLineSer = ctLineChart.getSerArray(seriesIndex);

            // Setting default data labels to show for line series, Note: all options default to true in case of "series.showLeaderLines(boolean)
            CTDLbls dLbls = ctLineSer.isSetDLbls() ? ctLineSer.getDLbls() : ctLineSer.addNewDLbls();
            ChartCustomizationHelper.setDataLabels(dLbls, false, false, false, false);

            // Setting shape and color
            CTMarker ctMarker = ctLineSer.isSetMarker() ? ctLineSer.getMarker() : ctLineSer.addNewMarker();
            ChartCustomizationHelper.configureSeriesCTShapeProperties(excelChartParameters, ctLineSer.addNewSpPr(),
                    ctMarker, seriesIndex);

            // using smooth line will also add ticks to the value axis.
            // i.e if some data point are at 0, 1 or 2 negative tick will also be added automatically to draw line smoothly
            CTBoolean ctBooleanSmooth = ctLineSer.isSetSmooth() ? ctLineSer.getSmooth() : ctLineSer.addNewSmooth();
            ctBooleanSmooth.setVal(excelChartParameters.isLineIsSmooth());


        }
    }

}
