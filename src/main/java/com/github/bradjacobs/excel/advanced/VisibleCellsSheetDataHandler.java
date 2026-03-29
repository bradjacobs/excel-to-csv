/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * Special SheetDataHandler that only processes 'visible' rows and columns.
 */
class VisibleCellsSheetDataHandler extends SheetDataHandler {

    private static final int NO_PREVIOUS_ROW = -1;

    /**
     * Tracks the last row index observed by {@link #startRow(int)}.
     */
    private int lastProcessedRowIndex = NO_PREVIOUS_ROW;

    private final SheetContext sheetContext;

    public VisibleCellsSheetDataHandler(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
            SheetContext sheetContext) {
        super(sheetConfig, stringRowConsumer);
        this.sheetContext = sheetContext;
    }

    @Override
    public void startRow(int rowNum) {
        int previousRowIndex = lastProcessedRowIndex;
        lastProcessedRowIndex = rowNum;
        fillMissingVisibleRows(previousRowIndex, rowNum);
    }

    /**
     * if configured to retain blank rows, then add in any
     * necessary missing blank rows between the previous
     * row that was processed and the current row.
     * @param previousRowIndex previous row index
     * @param currentRowIndex current row index
     */
    private void fillMissingVisibleRows(int previousRowIndex, int currentRowIndex) {
        if (shouldRetainBlankRows()) {
            int firstMissingRowIndex = previousRowIndex + 1;
            for (int rowIndex = firstMissingRowIndex; rowIndex < currentRowIndex; rowIndex++) {
                if (isRowVisible(rowIndex)) {
                    appendEmptyRow(); // add the blank row
                }
            }
        }
    }

    @Override
    public void cell(int rowNum, int columnIndex, String formattedValue, XSSFComment comment) {
        if (!shouldEmitCell(rowNum, columnIndex)) {
            return;
        }
        super.cell(rowNum, columnIndex, formattedValue, comment);
    }

    @Override
    protected void appendMissingColumnsBefore(int columnIndex) {
        // fill in any blanks between values in a row (if necessary)
        for (int col = getCurrentRowSize(); col < columnIndex; col++) {
            if (isColumnVisible(col)) {
                appendEmptyCellValue();
            }
        }
    }

    @Override
    protected void emitCurrentRow(int rowNum) {
        // only emit a row if it's visible
        if (isRowVisible(rowNum)) {
            super.emitCurrentRow(rowNum);
        }
    }

    private boolean shouldEmitCell(int rowNum, int columnIndex) {
        return isRowVisible(rowNum) && isColumnVisible(columnIndex);
    }

    private boolean isRowVisible(int rowNum) {
        return !sheetContext.isRowHidden(rowNum);
    }

    private boolean isColumnVisible(int columnIndex) {
        return !sheetContext.isColumnHidden(columnIndex);
    }
}