package com.anshuman.util;

import com.anshuman.exception.FormulaEvaluationException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

public class POIUtil {

    private POIUtil() {
        // use class statically
    }

    /**
     * Parse the cell-value as per the provided cell-type.
     *
     * @param cellValue A CellValue which contains a value
     * @param cellType  A CellType which contains the format/type of the cell
     * @return an Optional with the string value
     */
    public static Optional<String> parseByCellType(CellValue cellValue, CellType cellType) {
        return switch (cellType) {
            case BOOLEAN -> Optional.of(cellValue.getBooleanValue() ? "true" : "false");
            case STRING -> Optional.ofNullable(cellValue.getStringValue());
            case _NONE -> {System.out.println("none"); yield Optional.empty();}
            case NUMERIC -> Optional.of(cellValue.getNumberValue()).map(String::valueOf);
            case BLANK -> {System.out.println("blank"); yield Optional.empty();}
            case ERROR -> {System.out.println("error"); yield Optional.empty();}
            case FORMULA -> {System.out.println("formula"); yield Optional.empty();}
        };
    }

    /**
     * Throw exception if the cell is null
     *
     * @param cell the given cell
     */
    public static void nullCellCheck(Cell cell) {
        if (cell == null)
            throw new FormulaEvaluationException(new NullPointerException("Cannot perform operations on the cell, as it is null"));
    }

    public static String getSupportedFunctions() {
        return String.join(", ", WorkbookEvaluator.getSupportedFunctionNames());
    }

    public static String getUnsupportedFunctions() {
        return String.join(", ", WorkbookEvaluator.getNotSupportedFunctionNames());
    }

    public static String checkInputAgainstUnsupportedFunctions(List<String> formulae) {
        List<String> unsupportedFunctions = new ArrayList<>();
        WorkbookEvaluator.getNotSupportedFunctionNames()
            .forEach(function -> formulae.forEach(formula -> {
                    if (formula.contains(function))
                        unsupportedFunctions.add(function);
                }
        ));
        return String.join(", ", unsupportedFunctions);
    }

    public static String getNames(Workbook workbook) {
        return Optional.ofNullable(workbook)
            .filter(Objects::nonNull)
            .map(Workbook::getAllNames)
            .stream()
            .flatMap(Collection::stream)
            .map(POIUtil::stringifyName)
            .collect(Collectors.joining(",\n"));
    }

    public static String stringifyName(Name name) {
        return "{name: " + name.getNameName() + ", cellAddress: " + name.getRefersToFormula() + ", formula: " + name.getComment() + "}";
    }

    // give the range of cells to reference by name. e.g: for sheet1, row 1 and column 0, reference will be: 'sheet1'!$A$2
    public static String getCellReferenceWithSheet(String sheetName, int columnNumber, int rowNumber) {
        return "'" + sheetName + "'" + "!" + "$" + getColumnLetter(columnNumber) + "$" + (rowNumber + 1);
    }

    public static String getCellReference(int columnNumber, int rowNumber) {
        return getColumnLetter(columnNumber) + (rowNumber + 1);
    }

    public static String getColumnLetter(int columnNumber) {
        return "" + (char) (columnNumber + 65);
    }

    public static String getCellReference(CellAddress cellAddress) {
        return getCellReference(cellAddress.getColumn(), cellAddress.getRow());
    }

    public static void validateRow(Sheet sheet, int rowNumber) {
        if (rowNumber < 0)
            throw new FormulaEvaluationException("Cannot create row due to invalid row number: " + rowNumber);
        if (sheet == null)
            throw new FormulaEvaluationException("Cannot create row as the sheet does not exist");
        if (sheet.getRow(rowNumber) != null)
            throw new FormulaEvaluationException("Cannot create row at rowNumber: " + rowNumber + " as it already exists");
    }

    public static void validateCell(Row row, int columnNumber) {
        if (row == null)
            throw new FormulaEvaluationException("Cannot create cell as the row does not exist");
        int maxColumnNumber = 1048576;
        if (columnNumber < 0 || columnNumber > maxColumnNumber)
            throw new FormulaEvaluationException("Cannot create cell due to invalid column number " + columnNumber);
        if (row.getCell(columnNumber) != null)
            throw new FormulaEvaluationException("Cannot create cell at columnNumber: " + columnNumber + " as it already exists");
    }

    public static void validateSheetName(Workbook workbook, String name) {
        String msg = "Illegal name for sheet. ";
        if (name == null || name.isEmpty() || name.isBlank())
            throw new FormulaEvaluationException(msg + "Sheet name cannot be null/empty/blank");

        int nameLength = 31;
        if (name.length() > nameLength)
            throw new FormulaEvaluationException(msg + "Sheet name cannot be longer than " + nameLength + " chars");

        boolean containsIllegalChars = Stream
            .of(new char[] {0x0000, 0x0003, ':', '\\', '*', '?', '/', '[', ']'})
            .anyMatch(ch -> name.contains(new String(ch)));
        if (containsIllegalChars)
            throw new FormulaEvaluationException(msg + "Sheet name contains illegal characters");

        if (name.startsWith("'") || name.endsWith("'"))
            throw new FormulaEvaluationException(msg + "Sheet name cannot start or end with single quote");

        if (workbook.getSheet(name) != null)
            throw new FormulaEvaluationException("Cannot create sheet with name: " + name + " as it already exists");
    }
}
