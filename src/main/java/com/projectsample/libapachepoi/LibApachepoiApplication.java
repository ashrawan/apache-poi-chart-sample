package com.projectsample.libapachepoi;

import com.projectsample.libapachepoi.chart.generators.ExcelChartProperties;
import com.projectsample.libapachepoi.chart.util.ExcelColumn;
import com.projectsample.libapachepoi.playground.additional.SampleDataDTO;
import com.projectsample.libapachepoi.runner.ExcelMainLibRunner;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

import java.util.*;

@SpringBootApplication
public class LibApachepoiApplication {

    @Autowired
    private ExcelMainLibRunner excelMainLibRunner;

    public static void main(String[] args) {
        SpringApplication.run(LibApachepoiApplication.class, args);
    }

    @Bean
    public CommandLineRunner commandLineRunner() {
        return args -> {
            List<SampleDataDTO> languagesAndStats = new ArrayList<>();

            Map<String, Double> mapPython = new HashMap<>();
            mapPython.put("General Score", 100.0);
            mapPython.put("Jobs Availability", 88.22);
            mapPython.put("Trending", 100.00);

            Map<String, Double> mapCAndCPP = new HashMap<>();
            mapCAndCPP.put("General Score", 90.90);
            mapCAndCPP.put("Jobs Availability", 48.50);
            mapCAndCPP.put("Trending", 55.20);

            Map<String, Double> mapJava = new HashMap<>();
            mapJava.put("General Score", 70.22);
            mapJava.put("Jobs Availability", 95.07);
            mapJava.put("Trending", 74.19);

            Map<String, Double> mapJS = new HashMap<>();
            mapJS.put("General Score", 40.48);
            mapJS.put("Jobs Availability", 71.18);
            mapJS.put("Trending", 60.17);

            // Adding data rows (i.e series data)
            languagesAndStats.add(new SampleDataDTO("Python", Collections.unmodifiableMap(mapPython)));
            languagesAndStats.add(new SampleDataDTO("C/C++", Collections.unmodifiableMap(mapCAndCPP)));
            languagesAndStats.add(new SampleDataDTO("Java", Collections.unmodifiableMap(mapJava)));
            languagesAndStats.add(new SampleDataDTO("JavaScript", Collections.unmodifiableMap(mapJS)));

            // Category
            List<ExcelColumn> categoryColumns = new ArrayList<>();
            categoryColumns.add(new ExcelColumn("Category", 0, 0, 1, 3));

            // Chart to show
            ExcelChartProperties excelChartProperties = new ExcelChartProperties();
            excelChartProperties.setChartTitle("Language Stats");
            excelChartProperties.setChartPosition(ExcelChartProperties.ExcelPosition.BOTTOM);

            ExcelChartProperties.ExcelChartParameters excelChartParameters = new ExcelChartProperties.ExcelChartParameters();
            excelChartParameters.setCategoryAxisTitle("Category");
            excelChartParameters.setValueAxisTitle("Percent Value");
            excelChartParameters.setType(ExcelChartProperties.ExcelChartTypes.BAR);
            excelChartParameters.setCategoryColumns(Arrays.asList("Category"));
            excelChartParameters.setDataRows(Arrays.asList("Python", "C/C++", "Java", "JavaScript"));
            excelChartProperties.setParams(Arrays.asList(excelChartParameters));

            excelMainLibRunner.exportAsExcel(categoryColumns, languagesAndStats, excelChartProperties, "default-sheet01");
        };
    }
}
