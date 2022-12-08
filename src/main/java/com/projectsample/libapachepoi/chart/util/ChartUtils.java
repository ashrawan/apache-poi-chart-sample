package com.projectsample.libapachepoi.chart.util;

import com.projectsample.libapachepoi.chart.generators.ChartConfigurer;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.main.ThemeDocument;
import org.springframework.core.io.ClassPathResource;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Optional;


@Slf4j
public class ChartUtils {

    private static final String DEFAULT_PRIMARY_SET = "DEFAULT_PRIMARY_SET";
    private static final String DEFAULT_SECONDARY_SET = "DEFAULT_SECONDARY_SET";
    private static final String NEGATIVE_RED_SET = "NEGATIVE_RED_SET";
    private static final String POSITIVE_GREEN_SET = "POSITIVE_GREEN_SET";


    public static final Map<String, List<String>> COLOR_SETS_MAP = Map.of(
            DEFAULT_PRIMARY_SET, Arrays.asList("00AEEF", "F98E2B", "13D0CA", "A3D55F", "EDE819", "31006F"),
            DEFAULT_SECONDARY_SET, Arrays.asList("31006F", "9579D3", "EC008C", "209FED", "00B050", "F98E2B"),
            NEGATIVE_RED_SET, Arrays.asList("C00000", "D00000", "E00000", "FF0000", "FF3C3C", "FF8B8B"),
            POSITIVE_GREEN_SET, Arrays.asList("00B050", "00E668", "00FE73", "4BFF9C", "61FFA8", "A3FFCD")
    );

    public static int getIndexForColumn(String colName, List<ExcelColumn> columns) {
        Optional<ExcelColumn> optional = columns.stream()
                .filter(x -> x.getColName().equalsIgnoreCase(colName.trim()))
                .findFirst();
        return optional.isPresent() ? optional.get().getColumnStart() : -1;
    }

    public static Optional<ExcelColumn> getExcelColumn(String colName, List<ExcelColumn> columns) {
        Optional<ExcelColumn> optionalExcelColumn = columns.stream()
                .filter(x -> x.getColName().equalsIgnoreCase(colName.trim()))
                .findFirst();
        return optionalExcelColumn;
    }

    public static String getColumnNameForColumn(String colName,
                                                List<ExcelColumn> columns) {
        Optional<ExcelColumn> optional = columns.stream()
                .filter(x -> x.getColName().equalsIgnoreCase(colName.trim()))
                .findFirst();
        return optional.isPresent() ? optional.get().getColName() : "";
    }

    // Retrieving ThemesTable from workbook. Use this method only after ensuring that theme been set on workbook
    public static ThemesTable getThemesTable(ChartConfigurer chartConfigurer) {
        StylesTable st = chartConfigurer.getXssfSheet().getWorkbook().getStylesSource();
        return st.getTheme();
    }

    public static void setThemeFromTemplateFile(SXSSFWorkbook workbook) {

        try (XSSFWorkbook workbookWithTheme = new XSSFWorkbook(
                new ClassPathResource("ExcelThemeTemplate.xlsx").getInputStream())) {

            // setting accessor to public for ThemesTable "theme" field
            Field themeField = ThemesTable.class.getDeclaredField("theme");
            themeField.setAccessible(true);

            // Retrieving themeDocument from our template workbook
            ThemesTable themesTableTemp = workbookWithTheme.getStylesSource().getTheme();
            ThemeDocument themeDocumentTemp = (ThemeDocument) themeField.get(themesTableTemp);

            // Setting themeDocument to our workbook
            StylesTable st = workbook.getXSSFWorkbook().getStylesSource();
            st.ensureThemesTable();
            ThemesTable themesTable = st.getTheme();
            ThemeDocument themeDocument = (ThemeDocument) themeField.get(themesTable);
            themeDocument.setTheme(themeDocumentTemp.getTheme());

        } catch (IOException | NoSuchFieldException | IllegalAccessException e) {
            log.warn("Error occurred while initializing theme on workbook, {} ", e.getMessage());
        }

    }

}
