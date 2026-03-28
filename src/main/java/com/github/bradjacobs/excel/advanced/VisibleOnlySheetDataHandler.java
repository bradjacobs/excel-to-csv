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

    /**
     * Tracks the last row index observed by {@link #startRow(int)}.
     */
    private int lastProcessedRowIndex = -1;

    public VisibleOnlySheetDataHandler(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
            SheetContext sheetContext) {
        super(sheetConfig,
                stringRowConsumer,
                createVisibilityPolicy(sheetContext));
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