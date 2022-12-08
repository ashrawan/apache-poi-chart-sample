package com.projectsample.libapachepoi.chart.chartmodel;

import com.projectsample.libapachepoi.chart.generators.ChartConfigurer;
import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import com.projectsample.libapachepoi.chart.generators.helper.ChartCreationHelper;
import com.projectsample.libapachepoi.chart.generators.helper.ChartCustomizationHelper;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;

public class ExcelScatterChart implements ExcelChart {

    @Override
    public void drawChart(ChartConfigurer chartConfigurer) {

        XDDFChartData chartData = ChartCreationHelper.createChartData(chartConfigurer);
        if (chartConfigurer.isUseCustomChartCreation()) {
            customScatterCategoryAndSeriesDataSource(chartConfigurer);
        }
        ChartCreationHelper.initAndPlotChart(chartConfigurer, chartData);
        setScatterChartStylesAndLabels(chartConfigurer);
    }

    private CTScatterChart getCTScatterChart(ChartConfigurer chartConfigurer) {
        // Get last created Scatter chart element
        CTPlotArea plotArea = chartConfigurer.getXssfChart().getCTChart().getPlotArea();
        int totalScatterChartElements = plotArea.sizeOfScatterChartArray();
        return plotArea.getScatterChartArray(totalScatterChartElements - 1);
    }

    private void customScatterCategoryAndSeriesDataSource(ChartConfigurer chartConfigurer) {

        CTScatterChart ctScatterChart = getCTScatterChart(chartConfigurer);
        CTBoolean ctBooleanVaryColor = ctScatterChart.isSetVaryColors() ? ctScatterChart.getVaryColors() : ctScatterChart.addNewVaryColors();
        ctBooleanVaryColor.setVal(true);

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();

        for (int i = 0; i < excelChartParameters.getDataRows().size(); i++) {
            CTScatterSer ctScatterSer = ctScatterChart.addNewSer();
            ctScatterSer.addNewIdx().setVal(chartConfigurer.getSeriesStartId() + i);

            // ======== DEFINE CATEGORY Data source ==========
            CTAxDataSource cttAxDataSource = ctScatterSer.addNewXVal();
            ChartCreationHelper.createCategoryCustomDataSource(chartConfigurer,
                    excelChartParameters.getCategoryColumns(),
                    cttAxDataSource);

            // ============ DEFINE EACH SERIES data source ===============
            CTSerTx ctSerTx = ctScatterSer.addNewTx();
            CTNumDataSource ctNumDataSource = ctScatterSer.addNewYVal();
            ChartCreationHelper.createEachSeriesCustomDataSource(chartConfigurer,
                    excelChartParameters.getDataRows().get(i),
                    ctSerTx, ctNumDataSource);

        }

    }

    private void setScatterChartStylesAndLabels(ChartConfigurer chartConfigurer) {
        CTScatterChart ctScatterChart = getCTScatterChart(chartConfigurer);

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();
        if(excelChartParameters.getScatterStyle() == null) {
            excelChartParameters.setScatterStyle(ExcelChartProperties.ExcelScatterStyle.SCATTER_ONLY);
        }

        for (int seriesIndex = 0; seriesIndex < ctScatterChart.getSerList().size(); seriesIndex++) {

            // Setting default data labels to show for pie series, Note: all options default to true in case of "series.showLeaderLines(boolean)
            CTScatterSer ctScatterSer = ctScatterChart.getSerArray(seriesIndex);
            CTDLbls dLbls = ctScatterSer.isSetDLbls() ? ctScatterSer.getDLbls() : ctScatterSer.addNewDLbls();
            ChartCustomizationHelper.setDataLabels(dLbls, false, false, false, false);

            // Setting shape and color
            CTShapeProperties ctScatterSerShapeProperties = ctScatterSer.addNewSpPr();
            CTMarker ctMarker = ctScatterSer.isSetMarker() ? ctScatterSer.getMarker() : ctScatterSer.addNewMarker();
            ChartCustomizationHelper.configureSeriesCTShapeProperties(excelChartParameters, ctScatterSerShapeProperties,
                    ctMarker, seriesIndex);
            if(excelChartParameters.getScatterStyle() == ExcelChartProperties.ExcelScatterStyle.SCATTER_ONLY) {
                ChartCustomizationHelper.setLineAndShapeNoFill(true, false, ctScatterSerShapeProperties);
            }

            // using smooth line will also add ticks to the value axis.
            // i.e if some data point are at 0, 1 or 2 negative tick will also be added automatically to draw line smoothly
            CTBoolean ctBooleanSmooth = ctScatterSer.isSetSmooth() ? ctScatterSer.getSmooth() : ctScatterSer.addNewSmooth();
            ctBooleanSmooth.setVal(excelChartParameters.isLineIsSmooth());

        }
    }


}
