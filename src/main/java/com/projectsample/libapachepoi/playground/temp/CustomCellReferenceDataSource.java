package com.projectsample.libapachepoi.playground.temp;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import java.util.List;

// TODO: Testing Some Customizations
public class CustomCellReferenceDataSource implements XDDFNumericalDataSource<Double> {

    private SXSSFSheet sheet;
    private List<CellReference> cellReferenceList;
    private final XSSFFormulaEvaluator evaluator;

    public CustomCellReferenceDataSource(SXSSFSheet sheet, List<CellReference> cellReferenceList) {
        this.sheet = sheet;
        this.cellReferenceList = cellReferenceList;
        this.evaluator = sheet.getDrawingPatriarch().getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
//        this.evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
    }

    private String formatCode;

    @Override
    public String getFormatCode() {
        return formatCode;
    }

    @Override
    public void setFormatCode(String formatCode) {
        this.formatCode = formatCode;
    }

    @Override
    public int getPointCount() {
        return cellReferenceList.size();
    }

    @Override
    public Double getPointAt(int index) {
        CellValue cellValue = getCellValueAt(index);
        if (cellValue != null && cellValue.getCellType() == CellType.NUMERIC) {
            return cellValue.getNumberValue();
        } else {
            return null;
        }
    }

    protected CellValue getCellValueAt(int index) {
        if (index < 0 || index >= cellReferenceList.size()) {
            throw new IndexOutOfBoundsException(
                    "Index must be between 0 and " + (cellReferenceList.size() - 1) + " (inclusive), given: " + index);
        }
        CellReference cellReference = cellReferenceList.get(index);
        Row row = sheet.getRow(cellReference.getRow());
        Cell cell = row.getCell(cellReference.getCol());
        return (row == null) ? null : evaluator.evaluate(cell);
    }

    @Override
    public boolean isCellRange() {
        return true;
    }

    @Override
    public boolean isReference() {
        return true;
    }

    @Override
    public boolean isNumeric() {
        return true;
    }

    @Override
    public int getColIndex() {
        // Get first column
        CellReference cellReference = cellReferenceList.get(0);
        return cellReference.getCol();
    }

    @Override
    public String getDataRangeReference() {
//        cellRangeAddress.formatAsString(sheet.getSheetName(), true)
//        if (dataRange == null) {
//            throw new UnsupportedOperationException("Literal data source can not be expressed by reference.");
//        } else {
//            return dataRange;
//        }
//        return "INDEX((B2*ROW(1:7))/ROW(1:7),)";
        return "CHOOSE({1,2,3,4,5},(Sheet1!B2),(Sheet1!B2),(Sheet1!B2),(Sheet1!B2),(Sheet1!B2),(Sheet1!B2),(Sheet1!B2),(Sheet1!B2))";
//        return "INDEX(B1,C1,F1)";
    }

//    public String formatCellReferenceListAsString(String sheetName, boolean useAbsoluteAddress) {
//        StringBuilder sb = new StringBuilder();
//        if (sheetName != null) {
//            sb.append(SheetNameFormatter.format(sheetName));
//            sb.append("!");
//        }
//        CellReference cellRefFrom = new CellReference(getFirstRow(), getFirstColumn(),
//                useAbsoluteAddress, useAbsoluteAddress);
//        CellReference cellRefTo = new CellReference(getLastRow(), getLastColumn(),
//                useAbsoluteAddress, useAbsoluteAddress);
//        sb.append(cellRefFrom.formatAsString());
//
//        //for a single-cell reference return A1 instead of A1:A1
//        //for full-column ranges or full-row ranges return A:A instead of A,
//        //and 1:1 instead of 1
//        if(!cellRefFrom.equals(cellRefTo)
//                || isFullColumnRange() || isFullRowRange()){
//            sb.append(':');
//            sb.append(cellRefTo.formatAsString());
//        }
//        return sb.toString();
//    }

}
