package com.projectsample.libapachepoi.playground.customsamples;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PieChart {

	public static void main(String[] args) throws IOException {
		pieChart();
	}

	public static void pieChart() throws FileNotFoundException, IOException {
		try (XSSFWorkbook wb = new XSSFWorkbook()) {

			XSSFSheet sheet = wb.createSheet("CountryPieChart");

			// Create row and put some cells in it. Rows and cells are 0 based.
			Row row = sheet.createRow((short) 0);

			Cell cell = row.createCell((short) 0);
			cell.setCellValue("Russia");

			cell = row.createCell((short) 1);
			cell.setCellValue("Canada");

			cell = row.createCell((short) 2);
			cell.setCellValue("USA");

			cell = row.createCell((short) 3);
			cell.setCellValue("China");

			cell = row.createCell((short) 4);
			cell.setCellValue("Brazil");

			cell = row.createCell((short) 5);
			cell.setCellValue("Australia");

			cell = row.createCell((short) 6);
			cell.setCellValue("India");

			row = sheet.createRow((short) 1);

			cell = row.createCell((short) 0);
			cell.setCellValue(17098242);

			cell = row.createCell((short) 1);
			cell.setCellValue(9984670);

			cell = row.createCell((short) 2);
			cell.setCellValue(9826675);

			cell = row.createCell((short) 3);
			cell.setCellValue(9596961);

			cell = row.createCell((short) 4);
			cell.setCellValue(8514877);

			cell = row.createCell((short) 5);
			cell.setCellValue(7741220);

			cell = row.createCell((short) 6);
			cell.setCellValue(3287263);

			XSSFDrawing drawing = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 4, 7, 20);

			XSSFChart chart = drawing.createChart(anchor);
			chart.setTitleText("Area-wise Top Seven Countries");
			chart.setTitleOverlay(false);

			XDDFChartLegend legend = chart.getOrAddLegend();
			legend.setPosition(LegendPosition.TOP_RIGHT);

			XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet,
					new CellRangeAddress(0, 0, 0, 6));

			XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
					new CellRangeAddress(1, 1, 0, 6));

			XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);// chart.createData(ChartTypes.PIE,
																				// null, null);
			data.setVaryColors(true);
			data.addSeries(countries, values);
			chart.plot(data);

			// Write output to an excel file
			try (FileOutputStream fileOut = new FileOutputStream("SAMPLE-pie-chart.xlsx")) {
				wb.write(fileOut);
			}
		}
	}

}
