/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * SpecialSheetHandler that only processes visible rows and columns.
 */
class VisibleOnlySheetDataHandler extends SheetDataHandler {

    private static final int NO_PREVIOUS_ROW = -1;

    /**
     * Tracks the last row index observed by {@link #startRow(int)}.
     */
    private int lastProcessedRowIndex = NO_PREVIOUS_ROW;

    private final SheetContext sheetContext;

    public VisibleOnlySheetDataHandler(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
            SheetContext sheetContext) {
        super(sheetConfig, stringRowConsumer);
        this.sheetContext = sheetContext;
    }

    @Override
    public void startRow(int rowNum) {
        final int previousRowIndex = this.lastProcessedRowIndex;
        this.lastProcessedRowIndex = rowNum;
        fillMissingVisibleRows(previousRowIndex, rowNum);
    }

    private void fillMissingVisibleRows(int previousRowIndex, int currentRowIndex) {
        if (shouldRetainBlankRows()) {
            final int firstMissingRowIndex = previousRowIndex + 1;
            for (int rowIndex = firstMissingRowIndex; rowIndex < currentRowIndex; rowIndex++) {
                if (isRowVisible(rowIndex)) {
                    stringRowConsumer.accept(null);
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
        for (int col = currentRowValues.size(); col < columnIndex; col++) {
            if (isColumnVisible(col)) {
                currentRowValues.add(EMPTY_CELL_VALUE);
            }
        }
    }

    @Override
    public void endRow(int rowNum) {
        if (isRowVisible(rowNum)) {
            stringRowConsumer.accept(currentRowValues);
        }
        clearCurrentRow();
    }

    private boolean isRowVisible(int rowNum) {
        return !sheetContext.isRowHidden(rowNum);
    }

    private boolean isColumnVisible(int columnIndex) {
        return !sheetContext.isColumnHidden(columnIndex);
    }

    private boolean shouldEmitCell(int rowNum, int columnIndex) {
        return isRowVisible(rowNum) && isColumnVisible(columnIndex);
    }
}