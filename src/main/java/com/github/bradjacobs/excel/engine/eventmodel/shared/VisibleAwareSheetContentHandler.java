/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.shared;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.row.StringRowConsumer;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * Special SheetContentHandler that only processes 'visible' rows and columns.
 */
public class VisibleAwareSheetContentHandler extends SheetContentHandler {

    private static final int NO_PREVIOUS_ROW = -1;

    /**
     * Tracks the last row index observed by {@link #startRow(int)}.
     */
    private int lastProcessedRowIndex = NO_PREVIOUS_ROW;

    private final SheetVisibilityTracker sheetVisibilityTracker;

    public VisibleAwareSheetContentHandler(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
            SheetVisibilityTracker sheetVisibilityTracker) {
        super(sheetConfig, stringRowConsumer);
        this.sheetVisibilityTracker = sheetVisibilityTracker;
    }

    @Override
    public void startRow(int rowIndex) {
        int previousRowIndex = lastProcessedRowIndex;
        lastProcessedRowIndex = rowIndex;
        appendMissingVisibleRowsBetween(previousRowIndex, rowIndex);
    }

    /**
     * if configured to retain blank rows, then add in any
     * necessary missing blank rows between the previous
     * row that was processed and the current row.
     * @param previousRowIndex previous row index
     * @param currentRowIndex current row index
     */
    private void appendMissingVisibleRowsBetween(int previousRowIndex, int currentRowIndex) {
        if (shouldIncludeBlankRows()) {
            int firstMissingRowIndex = previousRowIndex + 1;
            for (int rowIndex = firstMissingRowIndex; rowIndex < currentRowIndex; rowIndex++) {
                if (isRowVisible(rowIndex)) {
                    appendEmptyRow();
                }
            }
        }
    }

    @Override
    public void cell(int rowIndex, int columnIndex, String formattedValue, XSSFComment comment) {
        if (!shouldEmitCell(rowIndex, columnIndex)) {
            return;
        }
        super.cell(rowIndex, columnIndex, formattedValue, comment);
    }

    /**
     * @inheritDoc
     */
    @Override
    protected void appendMissingColumnsBefore(int columnIndex) {
        // TODO - add tests for this case as it looks like it could be a bug
        for (int columnToFill = getCurrentRowSize(); columnToFill < columnIndex; columnToFill++) {
            if (isColumnVisible(columnToFill)) {
                appendEmptyCellValue();
            }
        }
    }

    @Override
    protected void emitCurrentRow(int rowIndex) {
        // only emit a row if it's visible
        if (isRowVisible(rowIndex)) {
            super.emitCurrentRow(rowIndex);
        }
    }

    private boolean shouldEmitCell(int rowIndex, int columnIndex) {
        return isRowVisible(rowIndex) && isColumnVisible(columnIndex);
    }

    private boolean isRowVisible(int rowIndex) {
        return sheetVisibilityTracker.isRowVisible(rowIndex);
    }

    private boolean isColumnVisible(int columnIndex) {
        return sheetVisibilityTracker.isColumnVisible(columnIndex);
    }
}

// Dev Note: some alternate implementations without inheritance may be possible
//   but all attempts have turned out to be "more ugly".
//   Also, could inject a type of "default policy" or "visible only policy" object
//   into SheetContentHandler, but that makes that class more convoluted.
