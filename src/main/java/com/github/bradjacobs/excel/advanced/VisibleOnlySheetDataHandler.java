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

    private final SheetVisibilityPolicy visibilityPolicy;

    public VisibleOnlySheetDataHandler(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
            SheetContext sheetContext) {
        super(sheetConfig, stringRowConsumer);
        this.visibilityPolicy = createVisibilityPolicy(sheetContext);
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

    private boolean shouldSkipCell(int rowNum, int columnIndex) {
        return !visibilityPolicy.isCellVisible(rowNum, columnIndex);
    }

    @Override
    public void cell(int rowNum, int columnIndex, String formattedValue, XSSFComment comment) {
        if (shouldSkipCell(rowNum, columnIndex)) {
            return;
        }
        super.cell(rowNum, columnIndex, formattedValue, comment);
    }

    @Override
    protected void appendMissingColumnsBefore(int columnIndex) {
        // fill in any blanks between values in a row (if necessary)
        for (int col = currentRowValues.size(); col < columnIndex; col++) {
            if (visibilityPolicy.isColumnVisible(col)) {
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
        return visibilityPolicy.isRowVisible(rowNum);
    }

    private static SheetVisibilityPolicy createVisibilityPolicy(SheetContext sheetContext) {
        return new SheetVisibilityPolicy() {
            @Override
            public boolean isRowVisible(int rowNum) {
                return !sheetContext.isRowHidden(rowNum);
            }

            @Override
            public boolean isColumnVisible(int columnIndex) {
                return !sheetContext.isColumnHidden(columnIndex);
            }
        };
    }
}