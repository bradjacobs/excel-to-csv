/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * Special SheetContentHandler that only processes 'visible' rows and columns.
 */
class VisibleCellsSheetContentHandler extends SheetContentHandler {

    private static final int NO_PREVIOUS_ROW = -1;

    /**
     * Tracks the last row index observed by {@link #startRow(int)}.
     */
    private int lastProcessedRowIndex = NO_PREVIOUS_ROW;

    private final SheetContext sheetContext;

    public VisibleCellsSheetContentHandler(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
            SheetContext sheetContext) {
        super(sheetConfig, stringRowConsumer);
        this.sheetContext = sheetContext;
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
        return sheetContext.isRowVisible(rowIndex);
    }

    private boolean isColumnVisible(int columnIndex) {
        return sheetContext.isColumnVisible(columnIndex);
    }
}