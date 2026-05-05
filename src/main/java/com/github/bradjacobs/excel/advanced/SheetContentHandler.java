/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.CellValueSanitizer;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.ArrayList;
import java.util.List;

// TODO: more javaDocs and unitTests
class SheetContentHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    private static final String EMPTY_CELL_VALUE = "";
    private static final String EXCEL_ERROR_PREFIX = "ERROR:";
    private static final String MISSING_CELL_REFERENCE_MESSAGE = "Unable to parse Excel Sheet. " +
            "A cell value was encountered without a cellReference.  " +
            "See 'Known Issues' for more details.";

    private final SheetConfig sheetConfig;
    private final CellValueSanitizer cellValueSanitizer;
    private final StringRowConsumer stringRowConsumer;
    private final List<String> currentRowValues = new ArrayList<>();

    public SheetContentHandler(SheetConfig sheetConfig, StringRowConsumer stringRowConsumer) {
        this.sheetConfig = sheetConfig;
        this.cellValueSanitizer = new CellValueSanitizer(
                sheetConfig.trimStringValues(),
                sheetConfig.getCharSanitizeFlags()
        );
        this.stringRowConsumer = stringRowConsumer;
    }

    @Override
    public void startRow(int rowNum) {
        appendMissingRowsBefore(rowNum);
    }

    @Override
    public void endRow(int rowNum) {
        emitCurrentRow(rowNum);
        clearCurrentRow();
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        CellAddress cellAddress = toCellAddress(cellReference);
        cell(cellAddress.getRow(), cellAddress.getColumn(), formattedValue, comment);
    }

    /**
     * A cell, with the given formatted value (may be null),
     * and possibly a comment (may be null), was encountered.
     */
    protected void cell(int rowNum, int columnIndex, String formattedValue, XSSFComment comment) {
        appendMissingColumnsBefore(columnIndex);
        appendCellValue(formattedValue);
    }

    protected void appendCellValue(String formattedValue) {
        currentRowValues.add(sanitizeCellValue(formattedValue));
    }

    /**
     * Appends an empty cell value to the current row.
     */
    protected void appendEmptyCellValue() {
        appendCellValue(EMPTY_CELL_VALUE);
    }

    private CellAddress toCellAddress(String cellReference) {
        Validate.isTrue(StringUtils.isNotEmpty(cellReference), MISSING_CELL_REFERENCE_MESSAGE);
        return new CellAddress(cellReference);
    }

    /**
     * Emits the current row to the consumer
     * @param rowNum current row number (for reference)
     */
    protected void emitCurrentRow(int rowNum) {
        stringRowConsumer.accept(currentRowValues);
    }

    // todo: better name?? (inverted)
    protected boolean shouldIncludeBlankRows() {
        return !sheetConfig.skipBlankRows();
    }

    protected void appendMissingRowsBefore(int rowNum) {
        if (shouldIncludeBlankRows()) {
            // add any filler blank rows (if necessary)
            while (stringRowConsumer.getRowCount() < rowNum) {
                appendEmptyRow();
            }
        }
    }

    /**
     * Adds a blank row to the output via the consumer.
     */
    protected void appendEmptyRow() {
        stringRowConsumer.accept(null);
    }

    /**
     * Returns the current row size (number of cells)
     * @return current row size
     */
    protected int getCurrentRowSize() {
        return currentRowValues.size();
    }

    protected void appendMissingColumnsBefore(int columnIndex) {
        // fill in any blanks between values in a row (if necessary)
        for (int col = currentRowValues.size(); col < columnIndex; col++) {
            appendEmptyCellValue();
        }
    }

    private String sanitizeCellValue(String cellValue) {
        return stripExcelErrorPrefix(cellValueSanitizer.sanitizeCellValue(cellValue));
    }

    /**
     * remove the first part of an error string to be consistent
     * with the behavior of reading cell values from Cell/Row/Sheet objects
     */
    private String stripExcelErrorPrefix(String input) {
        if (input.startsWith(EXCEL_ERROR_PREFIX)) {
            return input.substring(EXCEL_ERROR_PREFIX.length());
        }
        return input;
    }

    private void clearCurrentRow() {
        currentRowValues.clear();
    }
}