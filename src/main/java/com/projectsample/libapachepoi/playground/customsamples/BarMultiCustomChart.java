package com.projectsample.libapachepoi.playground.customsamples;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.io.FileOutputStream;
import java.util.Random;

public final class BarMultiCustomChart {

    private static final int NUM_OF_ROWS = 8;
    private static final Random RNG = new Random();

    public static void generateChart() throws Exception {
        try (SXSSFWorkbook wb = new SXSSFWorkbook(null, 100, true)) {

            SXSSFSheet sxssfSheet = wb.createSheet("Ranked Missions Completion");

            SXSSFRow row = sxssfSheet.createRow(0);
            row.createCell(0);
            row.createCell(1);
            row.createCell(2).setCellValue("Jiraiya");
            row.createCell(3).setCellValue("Kakashi Hatake");
            row.createCell(4).setCellValue("Minato Namikaze");


            SXSSFCell cell;
            int totalCategoryDistribution = 2;
            int initialDistSeq = 1;
            for (int r = 1; r <= NUM_OF_ROWS; r++) {
                row = sxssfSheet.createRow(r);

                if ((r + 1) % totalCategoryDistribution == 0) {
                    cell = row.createCell(0);
                    String rank = (r + 1) / 2 == 0 ? "Rank S|A" : "Other Rank";
                    cell.setCellValue(rank);
                    initialDistSeq = 1;
                }
                cell = row.createCell(1);
                cell.setCellValue("P" + initialDistSeq);

                cell = row.createCell(2);
                cell.setCellValue((int) RNG.nextInt(1500));
                cell = row.createCell(3);
                cell.setCellValue((int) RNG.nextInt(1200));
                cell = row.createCell(4);
                cell.setCellValue((int) RNG.nextInt(1000));

                initialDistSeq++;
            }

            sxssfSheet.createDrawingPatriarch();
            XSSFDrawing drawing = sxssfSheet.getDrawingPatriarch();
            ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 6, 0, (12 + 6), 20);
            XSSFSheet sheet = drawing.getSheet();

//            XSSFDrawing drawing = sheet.createDrawingPatriarch();
//            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 6, 0, (12 + 6), 20);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Ranked Missions Completion");
            chart.setTitleOverlay(false);

            //do not auto delete the title; is necessary for showing title in Calc
//            if (chart.getCTChart().getAutoTitleDeleted() == null) chart.getCTChart().addNewAutoTitleDeleted();
//            chart.getCTChart().getAutoTitleDeleted().setVal(false);

            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.RIGHT);

            CTChart ctChart = chart.getCTChart();
            CTPlotArea ctPlotArea = ctChart.getPlotArea();
            CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
            CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
            ctBoolean.setVal(true);

            //telling the BarChart that it has axes and giving them Ids
            ctBarChart.addNewAxId().setVal(123456);
            ctBarChart.addNewAxId().setVal(123457);

            //cat axis
            CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
            ctCatAx.addNewAxId().setVal(123456); //id of the cat axis
            CTScaling ctScaling = ctCatAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctCatAx.addNewDelete().setVal(false);
            ctCatAx.addNewAxPos().setVal(STAxPos.B);
            ctCatAx.addNewCrossAx().setVal(123457); //id of the val axis
            ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            //telling the category axis that it not has no multi level labels ;-)
            ctCatAx.addNewNoMultiLvlLbl().setVal(false);

            //val axis
            CTValAx ctValAx = ctPlotArea.addNewValAx();
            ctValAx.addNewAxId().setVal(123457); //id of the val axis
            ctScaling = ctValAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctValAx.addNewDelete().setVal(false);
            ctValAx.addNewAxPos().setVal(STAxPos.L);
            ctValAx.addNewCrossAx().setVal(123456); //id of the cat axis
            ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            // values axis Num Format
            ctValAx.addNewNumFmt().setSourceLinked(false);
            ctValAx.getNumFmt().setFormatCode("#0.00K");

            CTDispUnits dispUnits = ctValAx.addNewDispUnits();
            dispUnits.addNewBuiltInUnit().setVal(STBuiltInUnit.THOUSANDS);
            dispUnits.addNewDispUnitsLbl();

            //series
//            byte[][] seriesColors = new byte[][] {
//                    new byte[]{(byte)255, 0, 0}, //red
//                    new byte[]{0, (byte)255, 0}, //green
//                    new byte[]{0, 0, (byte)255}  //blue
//            };
            int seriesLength = 3;
            for (int i = 0; i < seriesLength; i++) {
                CTBarSer ctBarSer = ctBarChart.addNewSer();

                // Set series text
                CTSerTx ctSerTx = ctBarSer.addNewTx();
                CTStrRef ctStrRef = ctSerTx.addNewStrRef();
                ctStrRef.setF(
                        new CellRangeAddress(0, 0, i + 2, i + 2)
                                .formatAsString(sheet.getSheetName(), true));
                ctBarSer.addNewIdx().setVal(i);


                // Add Category Data source
                CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
                //do using MultiLvlStrRef instead of StrRef
                CTMultiLvlStrRef ctMultiLvlStrRef = cttAxDataSource.addNewMultiLvlStrRef();
                ctMultiLvlStrRef.setF(
                        new CellRangeAddress(1, NUM_OF_ROWS, 0, 1)
                                .formatAsString(sheet.getSheetName(), true));

                // Add value data source
                CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
                CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
                ctNumRef.setF(
                        new CellRangeAddress(1, NUM_OF_ROWS, i + 2, i + 2)
                                .formatAsString(sheet.getSheetName(), true));
//                CTNumData ctNumData = ctNumRef.addNewNumCache();
////                CTNumData ctNumData = ctNumDataSource.addNewNumLit();
//                ctNumData.addNewPtCount().setVal(1);
//                CTNumVal ctNumVal = ctNumData.addNewPt();
//                ctNumVal.setIdx(0);
//                ctNumVal.setV("" + 1500);

//                ctBarSer.addNewSpPr().addNewPattFill().addNewSrgbClr().setVal(seriesColors[i]);

//                CTColor ctColor = CTColor.Factory.newInstance();
//                CTSRgbColor srgbClr = ctColor.addNewSrgbClr();
//                srgbClr.setVal(ExcelTheme.hexToRGBByte("f55951"));
////                CTPatternFillProperties ctPatternFillProperties = ctBarSer.addNewSpPr().addNewPattFill();
////                ctPatternFillProperties.setPrst(STPresetPatternVal.WD_DN_DIAG);
////                ctPatternFillProperties.addNewFgClr().setSrgbClr(srgbClr);
//                CTSolidColorFillProperties ctSolidColorFillProperties = ctBarSer.addNewSpPr().addNewSolidFill();
//                ctSolidColorFillProperties.setSrgbClr(srgbClr);


                // Setting default data labels to show for bar series, Note: all options default to true in case of "series.showLeaderLines(boolean)
                CTDLbls dLbls = ctBarSer.isSetDLbls() ? ctBarSer.getDLbls() : ctBarSer.addNewDLbls();
                setDataLabels(dLbls, true, false, false, false);
                CTNumFmt ctNumFmt = dLbls.isSetNumFmt() ? dLbls.getNumFmt() : dLbls.addNewNumFmt();
//                ctNumFmt.setFormatCode("[>=1000000] $#,##0.0,,\"M\";[<1000000] $#,##0.0,\"K\";General");
                ctNumFmt.setFormatCode("#0.00K");
                ctNumFmt.setSourceLinked(false);
                dLbls.setNumFmt(ctNumFmt);

            }

            ctBarChart.addNewBarDir().setVal(STBarDir.COL);
            ctBarChart.addNewGrouping().setVal(STBarGrouping.CLUSTERED);

            STBarGrouping.Enum barGrouping = ctBarChart.getGrouping().getVal();

//            if (STBarGrouping.STACKED.equals(barGrouping) || STBarGrouping.PERCENT_STACKED.equals(barGrouping)) {
//                // Do this, only if its STACKED bar chart, correcting the "series overlap" so bars really are stacked and aligned properly
//                chart.getCTChart().getPlotArea().getBarChartArray(0).addNewOverlap().setVal((byte) 100);
//            } else {
//                // Sets excel "series overlap" option, which put some gap between bars
//                chart.getCTChart().getPlotArea().getBarChartArray(0).addNewOverlap().setVal((byte) -10);
//            }

            try (FileOutputStream fileOut = new FileOutputStream("SAMPLE-bar-multi-category-chart.xlsx")) {
                System.out.println("created - \"SAMPLE-bar-multi-category-chart.xlsx\"");
                wb.write(fileOut);
            }
        }
    }

    public static void setDataLabels(CTDLbls dLbls, boolean showValue, boolean showLegendKey, boolean showCategoryName, boolean showSeriesName) {
        dLbls.addNewShowVal().setVal(showValue);
        dLbls.addNewShowLegendKey().setVal(showLegendKey);
        dLbls.addNewShowCatName().setVal(showCategoryName);
        dLbls.addNewShowSerName().setVal(showSeriesName);
        dLbls.addNewShowPercent().setVal(false);
        dLbls.addNewShowBubbleSize().setVal(false);
    }
}
