package com.projectsample.libapachepoi.chart.generators.helper;

import com.projectsample.libapachepoi.chart.ExcelIndexDTO;
import com.projectsample.libapachepoi.chart.generators.ChartConfigurer;
import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import com.projectsample.libapachepoi.chart.util.ExcelColumn;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.springframework.util.StringUtils;

import java.util.Arrays;
import java.util.List;

import static com.projectsample.libapachepoi.chart.util.ChartUtils.*;

public class ChartCreationHelper {

    public static XDDFDataSource<String> createCategoryXDDFDataSource(ChartConfigurer chartConfigurer, String categoryColumnName) {
        ExcelIndexDTO chartIndexDTO = chartConfigurer.getChartIndexDTO();
        int colIndex = getIndexForColumn(categoryColumnName, chartConfigurer.getColumns());
        CellRangeAddress cellAddresses = new CellRangeAddress(
                chartIndexDTO.getDataRowStartIndex(), chartIndexDTO.getDataRowEndIndex(), colIndex, colIndex);
        return XDDFDataSourcesFactory
                .fromStringCellRange(chartConfigurer.getXssfSheet(), cellAddresses);
    }

    private static void createEachSeriesXDDFDataSource(ChartConfigurer chartConfigurer, String dataCol, XDDFChartData xddfChartData, XDDFDataSource<?> categoryDataSourceValues) {

        ExcelIndexDTO chartIndexDTO = chartConfigurer.getChartIndexDTO();
        int colIndex = getIndexForColumn(dataCol, chartConfigurer.getColumns());
        CellRangeAddress cellAddresses = new CellRangeAddress(
                chartIndexDTO.getDataRowStartIndex(), chartIndexDTO.getDataRowEndIndex(), colIndex, colIndex);
        XDDFNumericalDataSource<Double> tempData = XDDFDataSourcesFactory
                .fromNumericCellRange(chartConfigurer.getXssfSheet(), cellAddresses);
        XDDFChartData.Series series = xddfChartData
                .addSeries(categoryDataSourceValues, tempData);
        String seriesTitle = getColumnNameForColumn(dataCol, chartConfigurer.getColumns());
        series.setTitle(seriesTitle, null);
    }

    public static void createCategoryCustomDataSource(ChartConfigurer chartConfigurer, List<String> categoryColumns, CTAxDataSource cttAxDataSource) {

        String sheetName = chartConfigurer.getXssfSheet().getSheetName();

        // If categoryColumns size is greater than 1, i.e MultiLevelCategory
        // take first and last column as multi-level-category-range, else default to single level category
        boolean isMultiLevelCategory = categoryColumns.size() > 1;

        ExcelColumn excelColumn = getExcelColumn(categoryColumns.get(0), chartConfigurer.getColumns()).orElse(null);
        int colIndexStart = excelColumn.getColumnStart();
        int colIndexEnd = excelColumn.getColumnEnd();

        if (isMultiLevelCategory) {
            int lastCategoryIndex = categoryColumns.size() - 1;
            colIndexEnd = getIndexForColumn(categoryColumns.get(lastCategoryIndex), chartConfigurer.getColumns());
        }
        CellRangeAddress categoryCellAddresses = new CellRangeAddress(
                excelColumn.getRowStart(), excelColumn.getRowEnd(), colIndexStart, colIndexEnd);

        if (isMultiLevelCategory) {
            CTMultiLvlStrRef ctMultiLvlStrRef = cttAxDataSource.addNewMultiLvlStrRef();
            ctMultiLvlStrRef.setF(categoryCellAddresses.formatAsString(sheetName, true));

            // This will check "Multi-level category labels" as true for the "Horizontal (category) Axis Options"
            XDDFCategoryAxis categoryAxis = chartConfigurer.getCategoryAxis();
            CTPlotArea plotArea = chartConfigurer.getXssfChart().getCTChart().getPlotArea();
            CTCatAx ctCatAx = Arrays.stream(plotArea.getCatAxArray())
                    .filter(catAx -> catAx.getAxId().getVal() == categoryAxis.getId())
                    .findFirst()
                    .orElse(null);
            if (ctCatAx != null) {
                CTBoolean ctBooleanMultiLevel = ctCatAx.isSetNoMultiLvlLbl() ?
                        ctCatAx.getNoMultiLvlLbl() : ctCatAx.addNewNoMultiLvlLbl();
                ctBooleanMultiLevel.setVal(false);
            }
        } else {
            CTStrRef ctStrRefCat = cttAxDataSource.addNewStrRef();
            ctStrRefCat.setF(categoryCellAddresses.formatAsString(sheetName, true));
        }
    }

    public static void createEachSeriesCustomDataSource(ChartConfigurer chartConfigurer, String dataCol, CTSerTx ctSerTx, CTNumDataSource ctNumDataSource) {

        ExcelIndexDTO chartIndexDTO = chartConfigurer.getChartIndexDTO();
        String sheetName = chartConfigurer.getXssfSheet().getSheetName();

        // Set Each Series text - Single cell reference
        ExcelColumn categoryColumn = chartConfigurer.getColumns().get(0);
        int dataRowStartIndex = chartIndexDTO.getDataRowStartIndex() + 1;

        String seriesTitle = getColumnNameForColumn(dataCol, chartConfigurer.getColumns());
        if (StringUtils.hasText(seriesTitle)) {
            ctSerTx.setV(seriesTitle);
        } else {
            CellRangeAddress seriesTextCellAddresses = new CellRangeAddress(
                    dataRowStartIndex, dataRowStartIndex, 0, 0);
            CTStrRef ctStrRef = ctSerTx.addNewStrRef();
            ctStrRef.setF(seriesTextCellAddresses.formatAsString(sheetName, true));
        }

        // Set Each series NumDataSource - range cell reference
        CellRangeAddress seriesCellAddresses = new CellRangeAddress(dataRowStartIndex, dataRowStartIndex, categoryColumn.getColumnStart(), categoryColumn.getColumnEnd());
        CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
        ctNumRef.setF(seriesCellAddresses.formatAsString(sheetName, true));

        // increment dataRow start and end index
        chartIndexDTO.setDataRowStartIndex(chartIndexDTO.getDataRowStartIndex() + 1);
        chartIndexDTO.setDataRowEndIndex(chartIndexDTO.getDataRowEndIndex() + 1);
    }

    // Note: In case of "isUseCustomChartCreation" TRUE , series data are not initialized into XDDFChartData object
    // Series is initialized via XDDFDataSource, In-case of custom initialization; same is obtained from custom low-level-implementation
    public static XDDFChartData createChartData(ChartConfigurer chartConfigurer) {
        ExcelChartProperties.ExcelChartTypes excelChartType = chartConfigurer.getExcelChartParameters().getType();
        ChartTypes xddfChartType = getXDDFChartType(excelChartType);
        XSSFChart xssfChart = chartConfigurer.getXssfChart();
        return xssfChart.createData(xddfChartType, chartConfigurer.getCategoryAxis(), chartConfigurer.getValueAxis());
    }

    private static ChartTypes getXDDFChartType(ExcelChartProperties.ExcelChartTypes excelChartType) {
        switch (excelChartType) {
            case LINE:
                return ChartTypes.LINE;
            case BAR:
                return ChartTypes.BAR;
            case PIE:
                return ChartTypes.PIE;
            case SCATTER:
                return ChartTypes.SCATTER;
            default:
                return null;
        }
    }

    public static void initAndPlotChart(ChartConfigurer chartConfigurer, XDDFChartData xddfChartData) {

        ExcelChartProperties.ExcelChartParameters excelChartParameters = chartConfigurer.getExcelChartParameters();

        if (!chartConfigurer.isUseCustomChartCreation()) {
            // ======== DEFINE CATEGORY Data source ==========
            XDDFDataSource<?> categoryXDDFDataSource = createCategoryXDDFDataSource(chartConfigurer, excelChartParameters.getCategoryColumns().get(0));

            // ============ DEFINE EACH SERIES data source ===============
            for (String dataCol : excelChartParameters.getDataRows()) {
                createEachSeriesXDDFDataSource(chartConfigurer, dataCol, xddfChartData, categoryXDDFDataSource);
            }

            xddfChartData.setVaryColors(true);
        }
        chartConfigurer.getXssfChart().plot(xddfChartData);
    }

}
