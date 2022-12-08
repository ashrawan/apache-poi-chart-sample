# Excel Data and Charts

### Dev ReadMe:  
A Sample code showcasing how to work with ApachePOI Excel file and smart charts.  
#### 1. Generic Chart module : Act as a generic chart library that plot charts with existing data values.
#### 2. Playground:  Testing out some classes and functionality for ApachePOI chart.
#### 3.. ExcelMainLibRunner : Shows a simple approach for using the chart module.

## Introduction
Apache POI is a java library for working with Microsoft Office binary and OOXML file formats. Currently it has support for two formats OOXML and OLE2.

- **OLE2**: Object Linking & Embedding (It’s Microsoft’s Compound Document format to work with Microsoft files such as XLS, DOC, PPT etc. It’s the legacy implementation based on the OLEObject that uses approach of linking and embedding documents and other objects.)
- **OOXML**: Office Open XML (also informally known as OOXML is a new standards based XML file format in Microsoft Office 2007 and 2008. The file formats are such as. XLSX, DOCX, PPTX etc.)  

Generally, Apache POI has classified office documents using various API’s on the basis of following convention:
- Spread Sheets: SS = H*SS*F + X*SS*F
- Word Processing: WP = H*WP*F + X*WP*F
- PowerPoint Presentations (Slideshow): SL = H*SL*F + X*SL*F

Naming's used by POI:

- Components named “**H??F**” are for reading or writing OLE2 binary formats. (known as HF - **H**orrible **F**ormat).
- Components named “**X??F**” are for reading or writing OpenOffice XML (OOXML) formats. (known as XF – **X**ML **F**ormat).
   - SXSSF (since 3.8-beta3) – is an API-compatible streaming extension of XSSF to be used when very large spreadsheets have to be produced, and heap space is limited. e.g. SXSSFWorkbook, SXSSFSheet. SXSSF achieves its low memory footprint by limiting access to the rows that are within a sliding window, while XSSF gives access to all rows in the document.

## Working with Excel in Java

*Dependency*:

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.2</version>
</dependency>
```

#### Important Classes:
- **HSSF** – It’s implementation of the Excel ’97(-2007) file format. e.g. HSSFWorkbook, HSSFSheet. (HSSF – Horrible Spreadsheet Format)
- **XSSF** – It’s implementation of the Excel 2007 OOXML (.xlsx) file format. e.g. XSSFWorkbook, XSSFSheet. (Open Office XML Spreadsheet Format).

We will be using XSSF interface and its implementation.

#### Important packages:
Most of the classes for working with simple excel are available in package “org.apache.poi.xssf.usermodel”.
Apache POI has also introduced new package *“org.apache.poi.xddf.usermodel”*.

This provides base classes, enums and standards for XSSF classes and chart implementation.

### Supported Charts in Excel and Apache POI
We can view the charts supported in excel under “Recommended Charts” section.

Currently POI only supports this type of Excel charts:

```java 
public enum ChartTypes {
   AREA,
   AREA3D,
   BAR,
   BAR3D,
   DOUGHNUT,
   LINE,
   LINE3D,
   PIE,
   PIE3D,
   RADAR,
   SCATTER,
   SURFACE,
   SURFACE3D
}
```

## Creating and Working with charts in POI
General steps:
1.	Create workbook and create sheet.
2.	Add some data (Create a row and put some cells in it – later this cell value range will be used for chart creation)
3.	Create Drawing Patriarch and define anchor position (on defined position - chart will be drawn)
4.	Steps – Chart creation
      - Add and define legend position
      - Create Axis and set Position
      -	Define Data Source x axis, y-axis values. (CellRangeAddress)
      -	Define Data Source values to be plotted on those axis. (CellRangeAddress)
      -	Create Chart – ChartType, ChartAxis, ValueAxis
      -	Add data series – axis values/category , values
      -	Plot chart

#### 1. Create workbook and create sheet
```java 
XSSFWorkbook wb = new XSSFWorkbook()

XSSFSheet sheet = wb.createSheet("example_bar_chart")
```

#### 2. Add some data
```java
final int NUM_OF_ROWS = 3;
final int NUM_OF_COLUMNS = 10;

Row row;
Cell cell;

for (int rowIndex = 0; rowIndex < NUM_OF_ROWS; rowIndex++) {
row = sheet.createRow((short) rowIndex);

	for (int colIndex = 0; colIndex < NUM_OF_COLUMNS; colIndex++) {
		cell = row.createCell((short) colIndex);
		cell.setCellValue(colIndex * (rowIndex + 1.0)); // some random values
	}
}
```
#### 3. Create Drawing Patriarch and define anchor position
   Creates a new client anchor and sets the top-left and bottom-right coordinates of the anchor by cell references and offsets.
   
```java
XSSFDrawing drawing = sheet.createDrawingPatriarch();
XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 15);
```

start: 0, 5 (col, row)    i.e A, 5  
end: 10, 15 (col, row) i.e J, 15


#### 4. Steps – Chart creation
`   XSSFChart chart = drawing.createChart(anchor); `

- Add and define legend position:
```java 
XDDFChartLegend legend = chart.getOrAddLegend();

legend.setPosition(LegendPosition.TOP_RIGHT);
```

- Create Axis and set Position:

```java
// Use a category axis for the bottom axis.

XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);

bottomAxis.setTitle("x");

XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);

leftAxis.setTitle("f(x)");

leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
// Define Data Source x axis / y-axis values. (CellRangeAddress):
XDDFDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(0, 0, 0, NUM_OF_COLUMNS - 1));
// Define Data Source values to be plotted on those axis. (CellRangeAddress):
XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, NUM_OF_COLUMNS - 1));

XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, NUM_OF_COLUMNS - 1));
// Create Chart – ChartType, ChartAxis, ValueAxis:
XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
// Add data series – axis values/category , values:
XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(xs, ys1);
series1.setTitle("2x", null);
// Plot chart:
chart.plot(data);
```

### Official Links
https://poi.apache.org/  ( Official Apache POI website )  
https://github.com/apache/poi ( Source Code – GitHub Mirror Repository )


### Document Architecture, Testing and Debugging

**1. Open Packaging Conventions (OPC) (ECMA-376 OpenXML)**
- https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
- https://docs.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview


**2. Exploring “.xslx” file contents ( Parts and Relationships)**
- http://officeopenxml.com/drwOverview.php

### Overview on POI Implementation 
![Overview on POI abstract Implementation and low-level OOXML Implementation
](docs/POIComp.png?raw=true)

### For DEV:
Go through the package `playground` to run and explore various implementations. 

package `chart`, `excel`, and `runner` includes framework for a dynamic chart and data creation.

## Feel Free to Contribute !
