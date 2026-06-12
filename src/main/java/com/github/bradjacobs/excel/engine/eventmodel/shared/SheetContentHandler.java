/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.shared;

import com.github.bradjacobs.excel.engine.row.StringRowConsumer;
import com.github.bradjacobs.excel.sanitize.CellValueSanitizer;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.ArrayList;
import java.util.List;

// TODO: more javaDocs and unitTests
public class SheetContentHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    private static final String EMPTY_CELL_VALUE = "";
    private static final String EXCEL_ERROR_PREFIX = "ERROR:";
    private static final String MISSING_CELL_REFERENCE_MESSAGE = "Unable to parse Excel Sheet. " +
            "A cell value was encountered without a cellReference.  " +
            "See 'Known Issues' for more details.";
    private static final int NO_PREVIOUS_ROW = -1;

    public interface SheetContentEmitPolicy {
        boolean shouldEmitRow(int rowIndex);
        boolean shouldEmitColumn(int columnIndex);
        default boolean shouldEmitCell(int rowIndex, int columnIndex) {
            return shouldEmitRow(rowIndex) && shouldEmitColumn(columnIndex);
        }
    }

    // default policy is to emit everything
    private static final SheetContentEmitPolicy DEFAULT_POLICY =
            new SheetContentEmitPolicy() {
                @Override public boolean shouldEmitRow(int rowIndex) { return true; }
                @Override public boolean shouldEmitColumn(int columnIndex) { return true; }
            };

    private final StringRowConsumer stringRowConsumer;
    private final CellValueSanitizer cellValueSanitizer;
    private final SheetContentEmitPolicy emitPolicy;
    private final List<String> currentRowValues = new ArrayList<>();

    // Tracks the last row index observed by startRow(int).
    private int lastProcessedRowIndex = NO_PREVIOUS_ROW;

    public SheetContentHandler(
            StringRowConsumer stringRowConsumer,
            CellValueSanitizer cellValueSanitizer) {
        this(stringRowConsumer,  cellValueSanitizer, null);
    }

    public SheetContentHandler(
            StringRowConsumer stringRowConsumer,
            CellValueSanitizer cellValueSanitizer,
            SheetContentEmitPolicy emitPolicy) {
        this.stringRowConsumer = stringRowConsumer;
        this.cellValueSanitizer = cellValueSanitizer;
        this.emitPolicy = emitPolicy != null ? emitPolicy : DEFAULT_POLICY;
    }

    @Override
    public void startRow(int rowIndex) {
        appendMissingRowsBetween(lastProcessedRowIndex, rowIndex);
        lastProcessedRowIndex = rowIndex;
    }

    @Override
    public void endRow(int rowIndex) {
        emitCurrentRow(rowIndex);
        clearCurrentRow();
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        CellAddress cellAddress = toCellAddress(cellReference);
        cell(cellAddress.getRow(), cellAddress.getColumn(), formattedValue, comment);
    }

    /**
     * A cell, with the given formatted value (which may be null),
     * and possibly a comment (also may be null), was encountered.
     */
    private void cell(int rowIndex, int columnIndex, String formattedValue, XSSFComment comment) {
        if (emitPolicy.shouldEmitCell(rowIndex, columnIndex)) {
            // todo add more tests for this appendMissingColumns scenario.
            appendMissingColumnsBefore(columnIndex);
            appendCellValue(formattedValue);
        }
    }

    private void appendCellValue(String formattedValue) {
        currentRowValues.add(sanitizeCellValue(formattedValue));
    }

    private CellAddress toCellAddress(String cellReference) {
        Validate.isTrue(StringUtils.isNotEmpty(cellReference), MISSING_CELL_REFERENCE_MESSAGE);
        return new CellAddress(cellReference);
    }

    /**
     * Emits the current row to the consumer
     * @param rowIndex current row number (for reference)
     */
    private void emitCurrentRow(int rowIndex) {
        if (emitPolicy.shouldEmitRow(rowIndex)) {
            stringRowConsumer.accept(currentRowValues);
        }
    }

    /**
     * Add in any necessary missing blank rows between
     * the previous row that was processed and the current row.
     * @param previousRowIndex previous row index
     * @param currentRowIndex current row index
     */
    private void appendMissingRowsBetween(int previousRowIndex, int currentRowIndex) {
        // add any filler blank rows (if necessary)
        int firstMissingRowIndex = previousRowIndex + 1;
        for (int rowIndex = firstMissingRowIndex; rowIndex < currentRowIndex; rowIndex++) {
            if (emitPolicy.shouldEmitRow(rowIndex)) {
                // Emits a blank row to the output
                stringRowConsumer.accept(null);
            }
        }
    }

    /**
     * Appends any missing columns (if necessary)
     * @param columnIndex fill columns up to this index
     */
    private void appendMissingColumnsBefore(int columnIndex) {
        for (int columnToFill = currentRowValues.size(); columnToFill < columnIndex; columnToFill++) {
            if (this.emitPolicy.shouldEmitColumn(columnToFill)) {
                appendCellValue(EMPTY_CELL_VALUE);
            }
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
