/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.shared;

import com.github.bradjacobs.excel.engine.rows.StringRowConsumer;
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
    private static final List<String> EMPTY_ROW = List.of();
    private static final String EXCEL_ERROR_PREFIX = "ERROR:";
    private static final String MISSING_CELL_REFERENCE_MESSAGE = "Unable to parse Excel Sheet. " +
            "A cell value was encountered without a cellReference.  " +
            "See 'Known Issues' for more details.";
    private static final int NO_PREVIOUS_ROW = -1;
    private static final int NO_PREVIOUS_COLUMN = -1;

    public interface SheetContentEmitPolicy {
        boolean shouldEmitRow(int rowIndex);
        boolean shouldEmitColumn(int columnIndex);
        default boolean shouldEmitCell(int rowIndex, int columnIndex) {
            return shouldEmitRow(rowIndex) && shouldEmitColumn(columnIndex);
        }
    }

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
    private int lastProcessedColumnIndex = NO_PREVIOUS_COLUMN;

    public SheetContentHandler(
            StringRowConsumer stringRowConsumer,
            CellValueSanitizer cellValueSanitizer) {
        this(stringRowConsumer, cellValueSanitizer, null);
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
        lastProcessedColumnIndex = NO_PREVIOUS_COLUMN;
    }

    @Override
    public void endRow(int rowIndex) {
        emitCurrentRowIfIncluded(rowIndex);
        resetCurrentRow();
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        CellAddress cellAddress = parseCellAddress(cellReference);
        int rowIndex = cellAddress.getRow();
        int columnIndex = cellAddress.getColumn();
        if (shouldEmitCell(rowIndex, columnIndex)) {
            appendMissingColumnsBetween(lastProcessedColumnIndex, columnIndex);
            appendCellValue(formattedValue);
        }
        lastProcessedColumnIndex = columnIndex;
    }

    private CellAddress parseCellAddress(String cellReference) {
        Validate.isTrue(StringUtils.isNotEmpty(cellReference), MISSING_CELL_REFERENCE_MESSAGE);
        return new CellAddress(cellReference);
    }

    private void emitRowIfIncluded(int rowIndex, List<String> rowValues) {
        if (shouldEmitRow(rowIndex)) {
            stringRowConsumer.accept(rowValues);
        }
    }

    /**
     * Emits the current row to the consumer.
     * @param rowIndex current row number
     */
    private void emitCurrentRowIfIncluded(int rowIndex) {
        emitRowIfIncluded(rowIndex, currentRowValues);
    }

    /**
     * Adds any necessary missing blank rows between the previous row
     * that was processed and the current row.
     * @param previousRowIndex previous row index
     * @param currentRowIndex current row index
     */
    private void appendMissingRowsBetween(int previousRowIndex, int currentRowIndex) {
        int firstMissingRowIndex = previousRowIndex + 1;
        for (int rowIndex = firstMissingRowIndex; rowIndex < currentRowIndex; rowIndex++) {
            emitRowIfIncluded(rowIndex, EMPTY_ROW);
        }
    }

    /**
     * Appends any missing columns before the given column index.
     * @param previousColumnIndex previous row index
     * @param currentColumnIndex fill columns up to this index
     */
    private void appendMissingColumnsBetween(int previousColumnIndex, int currentColumnIndex) {
        int firstMissingColumnIndex = previousColumnIndex + 1;
        for (int columnIndex = firstMissingColumnIndex; columnIndex < currentColumnIndex; columnIndex++) {
            if (shouldEmitColumn(columnIndex)) {
                appendCellValue(EMPTY_CELL_VALUE);
            }
        }
    }

    private void appendCellValue(String formattedValue) {
        currentRowValues.add(sanitizeCellValue(formattedValue));
    }

    private String sanitizeCellValue(String cellValue) {
        return stripExcelErrorPrefix(cellValueSanitizer.sanitizeCellValue(cellValue));
    }

    /**
     * Removes the first part of an error string to be consistent with the
     * behavior of reading cell values from Cell/Row/Sheet objects.
     */
    private String stripExcelErrorPrefix(String input) {
        if (input.startsWith(EXCEL_ERROR_PREFIX)) {
            return input.substring(EXCEL_ERROR_PREFIX.length());
        }
        return input;
    }

    private boolean shouldEmitCell(int rowIndex, int columnIndex) {
        return emitPolicy.shouldEmitCell(rowIndex, columnIndex);
    }

    private boolean shouldEmitRow(int rowIndex) {
        return emitPolicy.shouldEmitRow(rowIndex);
    }

    private boolean shouldEmitColumn(int columnIndex) {
        return emitPolicy.shouldEmitColumn(columnIndex);
    }

    private void resetCurrentRow() {
        currentRowValues.clear();
    }
}
