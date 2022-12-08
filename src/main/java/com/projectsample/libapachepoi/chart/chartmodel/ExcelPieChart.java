package com.projectsample.libapachepoi.chart.chartmodel;

import com.projectsample.libapachepoi.chart.generators.ChartConfigurer;
import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import com.projectsample.libapachepoi.chart.generators.helper.ChartCreationHelper;
import com.projectsample.libapachepoi.chart.generators.helper.ChartCustomizationHelper;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

public class ExcelPieChart implements ExcelChart {

    @Override
    public void drawChart(ChartConfigurer chartConfigurer) {
        if (chartConfigurer.isUseCustomChartCreation()) {
            customPieCategoryAndSeriesDataSource(chartConfigurer);
        }
        XDDFChartData chartData = ChartCreationHelper.createChartData(chartConfigurer);
        ChartCreationHelper.initAndPlotChart(chartConfigurer, chartData);

        setPieChartStylesAndLabels(chartConfigurer, chartData);
    }

    private CTPieChart getCTPieChart(ChartConfigurer chartConfigurer) {
        // Get last created pie chart element
        CTPlotArea plotArea = chartConfigurer.getXssfChart().getCTChart().getPlotArea();
        int totalPieChartElements = plotArea.sizeOfPieChartArray();
        return plotArea.getPieChartArray(totalPieChartElements - 1);
    }

    private void customPieCategoryAndSeriesDataSource(ChartConfigurer chartConfigurer) {

        CTPieChart ctPieChart = getCTPieChart(chartConfigurer);
        CTBoolean ctBooleanVaryColor = ctPieChart.isSetVaryColors() ? ctPieChart.getVaryColors() : ctPieChart.addNewVaryColors();
        ctBooleanVaryColor.setVal(true);

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();

        for (int i = 0; i < excelChartParameters.getDataRows().size(); i++) {
            CTPieSer ctPieSer = ctPieChart.addNewSer();
            ctPieSer.addNewIdx().setVal(chartConfigurer.getSeriesStartId() + i);

            // ======== DEFINE CATEGORY Data source ==========
            CTAxDataSource cttAxDataSource = ctPieSer.addNewCat();
            ChartCreationHelper.createCategoryCustomDataSource(chartConfigurer,
                    excelChartParameters.getCategoryColumns(),
                    cttAxDataSource);

            // ============ DEFINE EACH SERIES data source ===============
            CTSerTx ctSerTx = ctPieSer.addNewTx();
            CTNumDataSource ctNumDataSource = ctPieSer.addNewVal();
            ChartCreationHelper.createEachSeriesCustomDataSource(chartConfigurer,
                    excelChartParameters.getDataRows().get(i),
                    ctSerTx, ctNumDataSource);

        }

    }

    private void setPieChartStylesAndLabels(ChartConfigurer chartConfigurer, XDDFChartData xddfChartData) {
        XDDFPieChartData xddfPieChartData = (XDDFPieChartData) xddfChartData;
        CTPieChart ctPieChart = getCTPieChart(chartConfigurer);

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();

        for (int seriesIndex = 0; seriesIndex < xddfPieChartData.getSeriesCount(); seriesIndex++) {

            // Setting default data labels to show for pie series, Note: all options default to true in case of "series.showLeaderLines(boolean)
            CTPieSer ctPieSer = ctPieChart.getSerArray(seriesIndex);
            CTDLbls dLbls = ctPieSer.isSetDLbls() ? ctPieSer.getDLbls() : ctPieSer.addNewDLbls();
            ChartCustomizationHelper.setDataLabels(dLbls, false, false, false, false);

            // Setting shape and color
            ChartCustomizationHelper.configureSeriesCTShapeProperties(excelChartParameters, ctPieSer.addNewSpPr(),
                    null, seriesIndex);

        }
    }

}
