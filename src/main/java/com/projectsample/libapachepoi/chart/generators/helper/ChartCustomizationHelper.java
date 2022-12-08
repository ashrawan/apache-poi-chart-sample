package com.projectsample.libapachepoi.chart.generators.helper;

import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import com.projectsample.libapachepoi.chart.util.ChartUtils;
import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.ObjectUtils;
import org.springframework.util.StringUtils;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import static com.projectsample.libapachepoi.chart.generators.ExcelChartProperties.ExcelChartParameters.SeriesStyleOptions;
import static com.projectsample.libapachepoi.chart.generators.ExcelChartProperties.ExcelChartTypes;
import static com.projectsample.libapachepoi.chart.generators.ExcelChartProperties.ExcelFillType;

public class ChartCustomizationHelper {

    private static final Logger LOGGER = LoggerFactory.getLogger(ChartCustomizationHelper.class);

    private ChartCustomizationHelper() {
    }

    /**
     * This method will customize shape and its properties for each series
     * If colorSet is supplied, will do cyclic iterate on colorSet and will retrieve color to use for particular series.
     * <p>
     * <p>
     * Note: if hexColor is passed into seriesStyleOptions this will take priority overriding the colorSet color.
     * Else, If either colorSet or hexColor isn't supplied, will return preemptively.
     * <p>
     *
     * @param excelChartParameters supplied chart params consisting chartType, fills, colorSet and additional properties
     * @param ctShapeProperties    ctChart series - specific shapeProperties object
     * @param ctMarker             pass ctChart series marker object, to configure marker as per supplied style properties
     * @param seriesIndex          ctChart current series index
     */
    public static void configureSeriesCTShapeProperties(ExcelChartProperties.ExcelChartParameters excelChartParameters,
                                                        CTShapeProperties ctShapeProperties, CTMarker ctMarker,
                                                        int seriesIndex) {

        // retrieving seriesStyleOptions, if supplied for the particular column
        String columnName = excelChartParameters.getDataRows().get(seriesIndex);
        Map<String, SeriesStyleOptions> seriesStyleOptionsMap =
                Optional.ofNullable(excelChartParameters.getSeriesStyleOptionsMap()).orElse(new HashMap<>());
        ExcelChartProperties.ExcelChartParameters.SeriesStyleOptions seriesStyleOptions = seriesStyleOptionsMap.get(columnName);

        String hexColor = null;
        if (StringUtils.hasText(excelChartParameters.getColorSet())) {
            String colorSetName = excelChartParameters.getColorSet();
            List<String> colorSets = ChartUtils.COLOR_SETS_MAP.get(colorSetName);
            int indexToRetrieve = seriesIndex < colorSets.size() ? seriesIndex : seriesIndex % (colorSets.size() - 1);
            hexColor = colorSets.get(indexToRetrieve);
        }
        if (seriesStyleOptions != null && StringUtils.hasText(seriesStyleOptions.getHexColor())) {
            hexColor = seriesStyleOptions.getHexColor();
        }

        ExcelChartTypes excelChartType = excelChartParameters.getType();
        if (seriesStyleOptions == null) {
            seriesStyleOptions = new SeriesStyleOptions();
            seriesStyleOptions.setFillType(ExcelFillType.SOLID);
            seriesStyleOptions.setHexColor(hexColor);
        }

        setFillProperties(excelChartType, ctShapeProperties, seriesStyleOptions);
        if (ctMarker != null) {
            CTShapeProperties ctMarkerShapeProperties = ctMarker.addNewSpPr();
            setFillProperties(excelChartType, ctMarkerShapeProperties, seriesStyleOptions);
            if (excelChartParameters.getScatterStyle() == ExcelChartProperties.ExcelScatterStyle.SCATTER_ONLY) {
                // For scatter chart, it always defaults to scatter-line
                // So, to show scatter only, noLineFill and default marker must be set
                if (ObjectUtils.isEmpty(seriesStyleOptions.getStMarkerStyle())) {
                    seriesStyleOptions.setStMarkerStyle(ExcelChartProperties.ExcelMarkerStyle.CIRCLE);
                }
                if(seriesStyleOptions.getMarkerSize() <= 0) {
                    seriesStyleOptions.setMarkerSize(8);
                }
            }
            configureCTMarkerProperties(ctMarker, seriesStyleOptions);
        }

    }

    /**
     * Sets MarkerStyle and its size
     *
     * @param ctMarker           complex type marker object added to specific ctChart series
     * @param seriesStyleOptions style options supplied for specific series column
     */
    public static void configureCTMarkerProperties(CTMarker ctMarker, SeriesStyleOptions seriesStyleOptions) {

        if (!ObjectUtils.isEmpty(seriesStyleOptions.getStMarkerStyle())) {
            CTMarkerStyle ctMarkerStyle = ctMarker.isSetSymbol() ? ctMarker.getSymbol() : ctMarker.addNewSymbol();
            STMarkerStyle.Enum stMarkerStyle = STMarkerStyle.Enum.forString(seriesStyleOptions.getStMarkerStyle().getValue());
            ctMarkerStyle.setVal(stMarkerStyle);
        }

        if (seriesStyleOptions.getMarkerSize() >= 0) {
            CTMarkerSize ctMarkerSize = ctMarker.isSetSize() ? ctMarker.getSize() : ctMarker.addNewSize();
            short markerSize = (short) (seriesStyleOptions.getMarkerSize() > 0 ? seriesStyleOptions.getMarkerSize() : 6);
            ctMarkerSize.setVal(markerSize);
        }
    }


    /**
     * Apply fill properties to the particular "ctShapeProperties" on the basis of supplied "seriesStyleOptions"
     * <p>
     * Note: If SRbgColor couldn't be formed, will return preemptively
     * On such case, excel will automatically defined color for that series, from theme palette
     *
     * @param excelChartType     chartType to evaluate additional properties while working with its shape
     * @param ctShapeProperties  ctChart series - specific shapeProperties object
     * @param seriesStyleOptions style options supplied for specific series column
     */
    public static void setFillProperties(ExcelChartTypes excelChartType, CTShapeProperties ctShapeProperties,
                                         SeriesStyleOptions seriesStyleOptions) {

        if (seriesStyleOptions.getHexColor() == null) {
            return;
        }
        byte[] bytesRBGColor = hexToRGBByte(seriesStyleOptions.getHexColor());
        if (bytesRBGColor.length <= 0) {
            return;
        }
        CTColor ctColor = CTColor.Factory.newInstance();
        CTSRgbColor srgbClr = ctColor.addNewSrgbClr();
        srgbClr.setVal(bytesRBGColor);

        switch (seriesStyleOptions.getFillType()) {
            case SOLID:
                CTSolidColorFillProperties ctSolidColorFillProperties = null;
                if (excelChartType.equals(ExcelChartTypes.LINE)) {
                    CTLineProperties ctLineProperties = ctShapeProperties.getLn() != null ? ctShapeProperties.getLn() : ctShapeProperties.addNewLn();
                    CTSolidColorFillProperties ctSolidLineColorFillProperties = ctLineProperties.isSetSolidFill() ?
                            ctLineProperties.getSolidFill() : ctLineProperties.addNewSolidFill();
                    ctSolidLineColorFillProperties.setSrgbClr(srgbClr);
                }
                ctSolidColorFillProperties = ctShapeProperties.isSetSolidFill() ?
                        ctShapeProperties.getSolidFill() : ctShapeProperties.addNewSolidFill();
                ctSolidColorFillProperties.setSrgbClr(srgbClr);
                break;
            case PATTERN:
                CTPatternFillProperties ctPatternFillProperties = ctShapeProperties.isSetPattFill() ?
                        ctShapeProperties.getPattFill() : ctShapeProperties.addNewPattFill();
                STPresetPatternVal.Enum stPresetPattern = STPresetPatternVal.WD_DN_DIAG;
                if (StringUtils.hasText(seriesStyleOptions.getStPresetPatternVal())) {
                    stPresetPattern = STPresetPatternVal.Enum.forString(seriesStyleOptions.getStPresetPatternVal());
                }
                ctPatternFillProperties.setPrst(stPresetPattern);
                CTColor ctForeGroundColor = ctPatternFillProperties.isSetFgClr() ?
                        ctPatternFillProperties.getFgClr() : ctPatternFillProperties.addNewFgClr();
                ctForeGroundColor.setSrgbClr(srgbClr);
                break;
            default:
                break;
        }
    }

    /**
     * Sets No fill properties for line and ct shape
     *
     * @param ctShapeProperties
     */
    public static void setLineAndShapeNoFill(boolean setLineNoFill, boolean setShapeNoFill, CTShapeProperties ctShapeProperties) {

        if (setLineNoFill) {
            CTLineProperties ctLineProperties = ctShapeProperties.getLn() != null ?
                    ctShapeProperties.getLn() : ctShapeProperties.addNewLn();
            CTNoFillProperties ctLineNoFillProperties = ctLineProperties.isSetNoFill() ?
                    ctLineProperties.getNoFill() : ctLineProperties.addNewNoFill();
            ctLineNoFillProperties.setNil();
        }

        if (setShapeNoFill) {
            CTNoFillProperties ctNoFillProperties = ctShapeProperties.isSetNoFill() ?
                    ctShapeProperties.getNoFill() : ctShapeProperties.addNewNoFill();
            ctNoFillProperties.setNil();
        }

    }

    /**
     * If a data point definition with the given index exists, then return it.
     * Otherwise create a new data point definition and return it.
     *
     * @param index data point index.
     * @return the CTDPt - data point with the given index.
     */
    public static CTDPt getOrCreateCTDataPoint(List<CTDPt> ctdPtList, long index) {
        for (int i = 0; i < ctdPtList.size(); i++) {
            if (ctdPtList.get(i).getIdx().getVal() == index) {
                return ctdPtList.get(i);
            }
            if (ctdPtList.get(i).getIdx().getVal() > index) {
                ctdPtList.add(i, CTDPt.Factory.newInstance());
                CTDPt point = ctdPtList.get(i);
                point.addNewIdx().setVal(index);
                return point;
            }
        }
        ctdPtList.add(CTDPt.Factory.newInstance());
        CTDPt point = ctdPtList.get(ctdPtList.size() - 1);
        point.addNewIdx().setVal(index);
        return point;
    }

    /**
     * Configure labels, if it needs to be displayed in each data series"
     *
     * @param dLbls
     * @param showValue        shows value at each data point
     * @param showLegendKey
     * @param showCategoryName
     * @param showSeriesName
     */
    public static void setDataLabels(CTDLbls dLbls, boolean showValue,
                                     boolean showLegendKey, boolean showCategoryName,
                                     boolean showSeriesName) {
        dLbls.addNewShowVal().setVal(showValue);
        dLbls.addNewShowLegendKey().setVal(showLegendKey);
        dLbls.addNewShowCatName().setVal(showCategoryName);
        dLbls.addNewShowSerName().setVal(showSeriesName);
        dLbls.addNewShowPercent().setVal(false);
        dLbls.addNewShowBubbleSize().setVal(false);
    }

    /**
     * Sets number format code for the data labels
     * format code is a excel formula/pattern, that changes display format for numeric numbers
     *
     * @param dLbls
     * @param formatCode
     */
    public static void setDataLabelNumFormat(CTDLbls dLbls, String formatCode) {
        CTNumFmt ctNumFmt = dLbls.isSetNumFmt() ? dLbls.getNumFmt() : dLbls.addNewNumFmt();
        ctNumFmt.setFormatCode(formatCode);
        ctNumFmt.setSourceLinked(false);
    }


    /**
     * Utility method, to convert hexColor string to rgb byte color
     *
     * @param rgbString hex color string e.g 00AEEF
     * @return rgb byte color
     */
    public static byte[] hexToRGBByte(String rgbString) {
        try {
            return Hex.decodeHex(rgbString);
        } catch (DecoderException e) {
            LOGGER.warn("Couldn't decode hex color string to rgb byte {} ", e.getMessage());
        }
        return new byte[0];
    }

}
