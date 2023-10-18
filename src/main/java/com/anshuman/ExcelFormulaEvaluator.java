package com.anshuman;

import static com.anshuman.NamedFormulaHelper.isFormula;
import static java.util.stream.Collectors.partitioningBy;
import static java.util.stream.Collectors.toMap;

import com.anshuman.exception.FormulaEvaluationException;
import com.anshuman.util.POIUtil;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Optional;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelFormulaEvaluator implements AutoCloseable {

    public static final String FILE_NAME_SUFFIX = ".xlsx";
    private static Workbook workbook;
    private final Path tempFilePath;
    private static FormulaEvaluator formulaEvaluator;
    private static ExcelFormulaEvaluator instance;

    public static Map<String, Boolean> validateFormulae(Map<String, String> formulaMap, String sheetName) {

        Map<String, Boolean> outputMap = new LinkedHashMap<>((int) (formulaMap.size() + 1.4));

        // create a map having two keys - true and false. Each key holds a subset of the formulaMap
        // key false holds only those key-value pairs of formula map that has constants.
        // key true holds only those key-value parts of formula map that has formulae.
        var map = formulaMap.entrySet()
            .stream()
            .collect(partitioningBy(entry -> isFormula(entry.getValue()),
                toMap(Entry::getKey, Entry::getValue)));

        // ExcelFormulaEvaluator uses "auto-closeable".
        // We use try-with resources pattern, so that the close method is automatically called.
        try (ExcelFormulaEvaluator instance = getInstance()) {
            // Initialize: start at address A1 in a new Excel sheet
            Sheet sheet = getOrCreateSheet(sheetName);
            final int[] rowNumber = {0};
            int columnNumber = 0;

            // first create names for only the pay elements with constant values.
            // These names are likely to be referenced in other formulae, so we want to process them first.
            // We don't have to evaluate the formulae for these. So directly get the value from the map.
            map.get(Boolean.FALSE)
                .forEach((key, value) -> outputMap.put(key,
                    instance.validateNamedConstant(sheet, key, value, rowNumber[0]++, columnNumber)));

            // then create names for only the pay elements with formulae (but not constants)
            // We're using a LinkedHashMap to save the sort order for the names.
            var formulaeToValidateMap = map.get(Boolean.TRUE)
                .entrySet()
                .stream()
                .sorted(NamedFormulaHelper::sortByDependency)
                .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue, (x, y) -> y, LinkedHashMap::new));

            formulaeToValidateMap
                .forEach((key, value) -> outputMap.put(key,
                    instance.validateNamedFormula(sheet, key, value, rowNumber[0]++, columnNumber)));

            System.out.println("Saved named formulae: " + POIUtil.getNames(workbook));

        } catch (Exception ex) {
            throw new FormulaEvaluationException("Exception encountered while calculating values for pay element formulae", ex);
        }

        return outputMap;
    }


    public static Map<String, Double> processFormulae(Map<String, String> formulaMap, String sheetName) {
        // We are using a linked hash map because we want to maintain the order of the pay elements.
        Map<String, Double> outputMap = new LinkedHashMap<>((int) (formulaMap.size() + 1.4));

        // create a map having two keys - true and false. Each key holds a subset of the formulaMap
        // key false holds only those key-value pairs of formula map that has constants.
        // key true holds only those key-value parts of formula map that has formulae.
        var map = formulaMap.entrySet()
            .stream()
            .collect(partitioningBy(entry -> isFormula(entry.getValue()),
                toMap(Entry::getKey, Entry::getValue)));

        // ExcelFormulaEvaluator uses "auto-closeable".
        // We use try-with resources pattern, so that the close method is automatically called.
        try (ExcelFormulaEvaluator instance = getInstance()) {
            // Initialize: start at address A1 in a new Excel sheet
            Sheet sheet = getOrCreateSheet(sheetName);
            final int[] rowNumber = {0};
            int columnNumber = 0;

            // first create names for only the pay elements with constant values.
            // These names are likely to be referenced in other formulae, so we want to process them first.
            // We don't have to evaluate the formulae for these. So directly get the value from the map.
            map.get(Boolean.FALSE)
                .forEach((key, value) -> outputMap.put(key,
                    instance.evaluateNamedConstant(sheet, key, value, rowNumber[0]++, columnNumber)));

            System.out.println("Saved named formulae: " + POIUtil.getNames(workbook));

            // then create names for only the pay elements with formulae (but not constants)
            // We're using a LinkedHashMap to save the sort order for the names.
            map.get(Boolean.TRUE)
                .entrySet()
                .stream()
                .sorted(NamedFormulaHelper::sortByDependency)
                .forEach(entry -> outputMap.put(entry.getKey(),
                instance.evaluateNamedFormula(sheet, entry.getKey(), entry.getValue(), rowNumber[0]++, columnNumber)));

            System.out.println("Saved named formulae: " + POIUtil.getNames(workbook));

        } catch (Exception ex) {
            throw new FormulaEvaluationException("Exception encountered while calculating values for pay element formulae", ex);
        }
        return outputMap;
    }

    // only allow a single instance for ExcelFormulaEvaluator to be created.
    private static ExcelFormulaEvaluator getInstance() {
        if (instance == null) {
            workbook = new HSSFWorkbook();
            formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            instance = new ExcelFormulaEvaluator();
        }
        return instance;
    }

    private ExcelFormulaEvaluator() {
        try {
            // create a temp Excel file with name format: temp-date-randomNumber.xlsx, on which we will perform operations
            String fileNamePrefix = "temp-" + LocalDateTime.now().format(DateTimeFormatter.ISO_DATE) + "-";
            tempFilePath = Files.createTempFile(fileNamePrefix, FILE_NAME_SUFFIX);
            System.out.println("Temp file created at: " + tempFilePath.toAbsolutePath());
            try (OutputStream fileOut = new FileOutputStream(tempFilePath.toFile())) {
                workbook.write(fileOut);
            }
        } catch (NullPointerException | IOException e) {
            throw new FormulaEvaluationException("Exception encountered while initializing workbook in the ExcelFormulaEvaluator constructor", e);
        }
    }

    private Sheet createSheet(String sheetName) {
        POIUtil.validateSheetName(workbook, sheetName);
        Sheet sheet = workbook.createSheet(sheetName);
        System.out.println("Created sheet with name: " + sheetName);
        return sheet;
    }

    private static Sheet getOrCreateSheet(String sheetName) {
        return (workbook.getSheet(sheetName) != null) ? workbook.getSheet(sheetName) : instance.createSheet(sheetName);
    }

    private Row createRow(Sheet sheet, int rowNumber) {
        POIUtil.validateRow(sheet, rowNumber);
        Row row = sheet.createRow(rowNumber);
        System.out.println("Created row: " + (row.getRowNum() + 1) + " on sheet: " + sheet.getSheetName());
        return row;
    }

    private static Row getOrCreateRow(Sheet sheet, int rowNumber) {
        return (sheet.getRow(rowNumber) == null) ? instance.createRow(sheet, rowNumber) : sheet.getRow(rowNumber);
    }

    private Cell createCell(Row row, int columnNumber, CellType cellType) {
        POIUtil.validateCell(row, columnNumber);
        Cell cell = (cellType != null) ? row.createCell(columnNumber, cellType) : row.createCell(columnNumber);
        System.out.println("Created cell at: " + POIUtil.getColumnLetter(cell.getColumnIndex()) +
            (cell.getRow().getRowNum() + 1) + ", with cellType: " + cell.getCellType().toString());
        return cell;
    }

    private static Cell getOrCreateCell(Row row, int columnNumber, CellType cellType) {
        return (row.getCell(columnNumber) == null) ? instance.createCell(row, columnNumber, cellType) : row.getCell(columnNumber);
    }

    private void setNumericCellValue(Cell cell, String name, double value) {
        cell.setCellValue(value);
        String cellAddress = POIUtil.getCellReference(cell.getAddress());
        System.out.println("cell: " + cellAddress + " with key: " + name + " set with value: " + value);
    }

    private void setCellFormula(Cell cell, String name, String formula) {
        POIUtil.nullCellCheck(cell);

        try {
            cell.setCellFormula(formula);
        } catch (FormulaParseException ex) {
            throw new FormulaEvaluationException("Cannot set formula: '" + formula +
                "', on cell: " + POIUtil.getCellReference(cell.getAddress()) +
                ", invalid formula or invalid syntax. Error Message: " + ex.getMessage(), ex);
        }

        // tell the cell that the formula for it has been updated
        // so that the cache doesn't store a pre-calculated value from a previous formula.
        formulaEvaluator.notifySetFormula(cell);
        String cellAddress = POIUtil.getCellReference(cell.getAddress().getColumn(), cell.getAddress().getRow());
        System.out.println("cell: " + cellAddress + " with key: " + name + " set with value: " + formula);
    }

    private Optional<String> evaluateFormula(Cell cell, CellType cellType) {
        POIUtil.nullCellCheck(cell);
        return Optional.of(cell)
            .map(ExcelFormulaEvaluator::evaluate)
            .flatMap(cellValue -> POIUtil.parseByCellType(cellValue, (cellType == null) ? cell.getCellType() : cellType))
            .filter(Predicate.not(String::isBlank));
    }

    private static CellValue evaluate(Cell cell) {
        try {
            return formulaEvaluator.evaluate(cell);
        } catch (NotImplementedException ex) {
            throw new FormulaEvaluationException("Formula " + cell.getCellFormula() +
                " at cell: " + POIUtil.getCellReference(cell.getAddress()) +
                " is not implemented in POI", ex);
        }
    }

    /**
     * Create a name for an area (range of cells) or cell. We are storing formulae for each pay element in a new cell and then naming that cell. This allows us
     * to refer to formula by name. It also enables a required feature, that other formulae can contain previously defined names.
     *
     * @param sheet        the sheet to which the cell belongs
     * @param nameName     the name of the cell
     * @param columnNumber the column number of the cell
     * @param rowNumber    the row number of the cell
     */
    private String createName(Sheet sheet, String nameName, String formula, int columnNumber, int rowNumber) {
        Name name = workbook.createName();
        name.setSheetIndex(workbook.getSheetIndex(sheet));
        name.setNameName(nameName);
        name.setComment(formula);
        try {
            name.setRefersToFormula(POIUtil.getCellReferenceWithSheet(sheet.getSheetName(), columnNumber, rowNumber));
        } catch (IllegalArgumentException ex) {
            throw new FormulaEvaluationException("Cannot create name, as the formula text cannot be parsed", ex);
        }
        System.out.println("Name created : " + POIUtil.stringifyName(name));
        return name.getNameName();
    }

    private double evaluateNamedFormula(Sheet sheet, String name, String value, int rowNumber, int columnNumber) {
        String nameName = createName(sheet, name, value, columnNumber, rowNumber);
        Row row = getOrCreateRow(sheet, rowNumber);
        Cell cell = getOrCreateCell(row, columnNumber, CellType.FORMULA);
        instance.setCellFormula(cell, nameName, value);
        return Double.parseDouble(instance.evaluateFormula(cell, CellType.NUMERIC).orElse(""));
    }

    private Boolean validateNamedFormula(Sheet sheet, String name, String value, int rowNumber, int columnNumber) {
        String nameName;
        try {
            nameName = createName(sheet, name, value, columnNumber, rowNumber);
        } catch (FormulaEvaluationException ex) {
            if (ex.getCause() instanceof IllegalArgumentException) {
                System.err.println(ex.getMessage());
                return false;
            }
            else throw ex;
        }
        Row row = getOrCreateRow(sheet, rowNumber);
        Cell cell = getOrCreateCell(row, columnNumber, CellType.FORMULA);
        try {
            instance.setCellFormula(cell, nameName, value);
        } catch (FormulaEvaluationException ex) {
            if (ex.getCause() instanceof FormulaParseException) {
                System.err.println(ex.getMessage());
                return false;
            }
            else throw ex;
        }
        String outputStr;
        try {
            outputStr = instance.evaluateFormula(cell, CellType.NUMERIC).orElse("");
        } catch (FormulaEvaluationException ex) {
            if (ex.getCause() instanceof NotImplementedException) {
                System.err.println(ex.getMessage());
                return false;
            }
            else throw ex;
        }
        try {
            Double.parseDouble(outputStr);
        } catch (NullPointerException | NumberFormatException ex) {
            System.err.println("Cannot parse output: " + outputStr + " to double, errorMsg: " + ex.getMessage());
            return false;
        }
        // if there are no exceptions, return true
        return true;
    }

    private double evaluateNamedConstant(Sheet sheet, String name, String value, int rowNumber, int columnNumber) {
        String nameName = createName(sheet, name, value, columnNumber, rowNumber);
        double doubleValue = Double.parseDouble(value);
        // create a new cell on each new row, with type numeric,
        // since the input value is a constant (number) and not a formula (string)
        Row row = getOrCreateRow(sheet, rowNumber);
        Cell cell = getOrCreateCell(row, columnNumber, CellType.NUMERIC);
        instance.setNumericCellValue(cell, nameName, doubleValue);
        return doubleValue;
    }

    private Boolean validateNamedConstant(Sheet sheet, String name, String value, int rowNumber, int columnNumber) {

        String nameName;
        // if formula cannot be parsed, return false
        try {
            nameName = createName(sheet, name, value, columnNumber, rowNumber);
        } catch (FormulaEvaluationException ex) {
            if (ex.getCause() instanceof IllegalArgumentException)
                return false;
            else throw ex;
        }

        double doubleValue;
        try {
            doubleValue = Double.parseDouble(value);
        } catch (NullPointerException | NumberFormatException ex) {
            System.err.println("Cannot parse output: " + value + " to double, errorMsg: " + ex.getMessage());
            return false;
        }

        // create a new cell on each new row, with type numeric,
        // since the input value is a constant (number) and not a formula (string)
        Row row = getOrCreateRow(sheet, rowNumber);
        Cell cell = getOrCreateCell(row, columnNumber, CellType.NUMERIC);
        instance.setNumericCellValue(cell, nameName, doubleValue);
        // if there are no exceptions, return true
        return true;
    }

    @Override
    public void close() throws Exception {
        String errors = closeWorkbook();
        if (tempFilePath != null)
            errors += deleteTempFile();
        if (!errors.isEmpty())
            System.out.println("Clean up completed with errors: " + errors);
        instance = null;
        System.out.println("performing cleanup: completed");
    }

    private String closeWorkbook() {
        try {
            workbook.close();
            System.out.println("performing cleanup: closed workbook");
        } catch (IOException ex) {
            return "Could not close workbook object: " +  ex.getMessage();
        }
        return "";
    }

    private String deleteTempFile() {
        String fileName = tempFilePath.getFileName().toString();
        try {
            if (Files.deleteIfExists(tempFilePath))
                System.out.println("performing cleanup: deleted temp excel file: " + fileName);
        } catch (IOException ex) {
            return " Could not delete temp file: " + ex.getMessage();
        }
        return "";
    }
}

class NamedFormulaHelper {

    // We want to sort the formulae so that the named formula is processed first before any formula that references it.
    // Between two key-value pairs, we compare the key (named formula) against the value (formula which may refer to a named formula)
    // to manage the ordering.
    public static int sortByDependency(Entry<String, String> e1, Entry<String, String> e2) {
        return (!e1.getValue().contains(e2.getKey())) ? -1: 1;
    }


    public static boolean isFormula(String value) {
        try {
            Double.parseDouble(value);
            return false;
        } catch (Exception ex) {
            return true;
        }
    }

    private NamedFormulaHelper() {
        // use class statically
    }
}
