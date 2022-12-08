package com.projectsample.libapachepoi.chart.generators;

import lombok.Data;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

@Data
public class ExcelChartProperties implements Serializable {
    private ExcelPosition chartPosition;
    private ExcelPosition legendPosition;
    private String chartTitle;
    private List<ExcelChartParameters> params;

    @Data
    public static class ExcelChartParameters implements Serializable {
        private ExcelChartTypes type;

        // display title
        private String categoryAxisTitle;
        private String valueAxisTitle;

        // Axis data, Category and Values columns
        private List<String> categoryColumns;
        private List<String> dataRows;

        // chart specific params
        private ExcelBarGrouping barGrouping;
        private ExcelBarDirection barDirection;
        private int barSeriesOverlapPercent;
        private ExcelScatterStyle scatterStyle;

        // Optional customizer params
        private boolean lineIsSmooth;
        private boolean plotOnSecondaryAxis;
        private boolean useSameMinMaxScaleAsPrimary;
        private String colorSet;
        private Map<String, SeriesStyleOptions> seriesStyleOptionsMap;

        @Data
        public static class SeriesStyleOptions implements Serializable {
            private ExcelFillType fillType;
            private String stPresetPatternVal;
            private String hexColor;

            private ExcelMarkerStyle stMarkerStyle;
            private int markerSize;

            private int transparencyPercent;
        }

    }

    public enum ExcelPosition {
        BOTTOM,
        LEFT,
        RIGHT,
        TOP,
        TOP_RIGHT;
    }

    public enum ExcelChartTypes {
        BAR,COLUMN,LINE,PIE,SCATTER,NONE
    }

    public enum ExcelBarDirection {
        BAR,
        COL;
    }

    public enum ExcelBarGrouping {
        STANDARD,
        CLUSTERED,
        STACKED,
        PERCENT_STACKED;
    }

    public enum ExcelScatterStyle {
        SCATTER_ONLY,
        LINE,
        LINE_MARKER,
        MARKER,
        SMOOTH,
        SMOOTH_MARKER
    }

    public enum ExcelMarkerStyle {
        CIRCLE("circle"),
        DASH("dash"),
        DIAMOND("diamond"),
        DOT("dot"),
        PLUS("plus"),
        SQUARE("square"),
        STAR("star"),
        TRIANGLE("triangle"),
        X("x"),
        AUTO("auto"),
        NONE("none");

        private String value;

        ExcelMarkerStyle(String value) {
            this.value = value;
        }

        public String getValue(){
            return value;
        }
    }

    public enum ExcelFillType {
        SOLID, PATTERN, NONE
    }

}
