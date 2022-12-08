package com.projectsample.libapachepoi.chart.chartmodel;


import com.projectsample.libapachepoi.chart.generators.ChartConfigurer;
import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import com.projectsample.libapachepoi.chart.generators.helper.ChartCreationHelper;
import com.projectsample.libapachepoi.chart.generators.helper.ChartCustomizationHelper;
import org.apache.poi.xddf.usermodel.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.util.List;

public class ExcelBarChart implements ExcelChart {

    @Override
    public void drawChart(ChartConfigurer chartConfigurer) {
        XDDFChartData chartData = ChartCreationHelper.createChartData(chartConfigurer);
        if (chartConfigurer.isUseCustomChartCreation()) {
            customBarCategoryAndSeriesDataSource(chartConfigurer);
        }
        ChartCreationHelper.initAndPlotChart(chartConfigurer, chartData);

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();
        setBarGroupingAndDirection(chartConfigurer, chartData, excelChartParameters.getBarDirection(),
                excelChartParameters.getBarGrouping());
        setBarChartSeriesStylesAndLabels(chartConfigurer);
    }

    private CTBarChart getCTBarChart(ChartConfigurer chartConfigurer) {
        // Get last created bar chart element
        CTPlotArea plotArea = chartConfigurer.getXssfChart().getCTChart().getPlotArea();
        int totalBarChartElements = plotArea.sizeOfBarChartArray();
        return plotArea.getBarChartArray(totalBarChartElements - 1);
    }

    private void customBarCategoryAndSeriesDataSource(ChartConfigurer chartConfigurer) {

        CTBarChart ctBarChart = getCTBarChart(chartConfigurer);
        CTBoolean ctBooleanVaryColor = ctBarChart.isSetVaryColors() ?
                ctBarChart.getVaryColors() : ctBarChart.addNewVaryColors();
        ctBooleanVaryColor.setVal(true);

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();
        for (int i = 0; i < excelChartParameters.getDataRows().size(); i++) {

            CTBarSer ctBarSer = ctBarChart.addNewSer();
            ctBarSer.addNewIdx().setVal(chartConfigurer.getSeriesStartId() + i);

            // ======== DEFINE CATEGORY Data source ==========
            CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
            ChartCreationHelper.createCategoryCustomDataSource(chartConfigurer,
                    excelChartParameters.getCategoryColumns(),
                    cttAxDataSource);

            // ============ DEFINE EACH SERIES data source ===============
            CTSerTx ctSerTx = ctBarSer.addNewTx();
            CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
            ChartCreationHelper.createEachSeriesCustomDataSource(chartConfigurer,
                    excelChartParameters.getDataRows().get(i),
                    ctSerTx, ctNumDataSource);

        }

    }

    private void setBarGroupingAndDirection(ChartConfigurer chartConfigurer, XDDFChartData xddfChartData,
                                            ExcelChartProperties.ExcelBarDirection excelBarDirection,
                                            ExcelChartProperties.ExcelBarGrouping excelBarGrouping) {

        BarDirection barDirection = excelBarDirection != null ?
                BarDirection.valueOf(excelBarDirection.name()) : BarDirection.COL;
        BarGrouping barGrouping = excelBarGrouping != null ?
                BarGrouping.valueOf(excelBarGrouping.name()) : BarGrouping.STANDARD;

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();
        List<XDDFValueAxis> valueAxes = xddfChartData.getValueAxes();
        for (XDDFValueAxis xddfValueAxis : valueAxes) {
            // BETWEEN category axis crosses the value axis between the strokes and not midpoint the strokes.
            // Else the bars are only half wide visible for first and last category.
            xddfValueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

            // In-case of PERCENT_STACKED axis start value is auto-adjusted, so setting start value as 0%.
            if (BarGrouping.PERCENT_STACKED.equals(barGrouping)) {
                xddfValueAxis.setMinimum(0.0);
            }
        }

        XDDFBarChartData xddfBarChartData = (XDDFBarChartData) xddfChartData;
        xddfBarChartData.setBarDirection(barDirection);
        xddfBarChartData.setBarGrouping(barGrouping);

        if (BarGrouping.STACKED.equals(barGrouping)
                || BarGrouping.PERCENT_STACKED.equals(barGrouping)) {
            // Do this, only if its STACKED bar chart, correcting the "series overlap"
            // So bars really are stacked and aligned properly
            CTBarChart ctBarChart = getCTBarChart(chartConfigurer);
            ctBarChart.addNewOverlap().setVal((byte) 100);

        } else {
            // Sets excel "series overlap" option, which put some gap between bars
            CTBarChart ctBarChart = getCTBarChart(chartConfigurer);
            ctBarChart.addNewOverlap().setVal(excelChartParameters.getBarSeriesOverlapPercent());
        }

    }

    private void setBarChartSeriesStylesAndLabels(ChartConfigurer chartConfigurer) {
        CTBarChart ctBarChart = getCTBarChart(chartConfigurer);

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();

        for (int seriesIndex = 0; seriesIndex < ctBarChart.getSerList().size(); seriesIndex++) {

            // Setting default data labels to show for bar series,
            // Note: all options default to true in case of "series.showLeaderLines(boolean)
            CTBarSer ctBarSer = getCTBarChart(chartConfigurer).getSerArray(seriesIndex);
            CTDLbls dLbls = ctBarSer.isSetDLbls() ? ctBarSer.getDLbls() : ctBarSer.addNewDLbls();
            ChartCustomizationHelper.setDataLabels(dLbls, false, false, false,
                    false);

            // Setting shape and color
            ChartCustomizationHelper.configureSeriesCTShapeProperties(excelChartParameters, ctBarSer.addNewSpPr(),
                    null, seriesIndex);

        }
    }

}
