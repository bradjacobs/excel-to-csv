package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.SheetConfig;
import com.github.bradjacobs.excel.StringRowConsumer;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * SpecialSheetHandler that only processes visible rows and columns.
 */
class VisibleOnlySpecialSheetHandler extends SpecialSheetHandler {

    private final SheetContext sheetContext;

    /**
     * Tracks the last row index observed by {@link #startRow(int)}.
     */
    private int lastProcessedRowIndex = -1;

    public VisibleOnlySpecialSheetHandler(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
            SheetContext sheetContext) {
        super(sheetConfig, stringRowConsumer);
        this.sheetContext = sheetContext;
    }

    private boolean isHiddenRow(int rowIndex) {
        return sheetContext.isRowHidden(rowIndex);
    }

    private boolean isHiddenColumn(int columnIndex) {
        return sheetContext.isColumnHidden(columnIndex);
    }

    @Override
    public void startRow(int rowNum) {
        final int previousRowIndex = this.lastProcessedRowIndex;
        this.lastProcessedRowIndex = rowNum;
        fillMissingVisibleRows(previousRowIndex, rowNum);
    }

    private void fillMissingVisibleRows(int previousRowIndex, int currentRowIndex) {
        if (sheetConfig.isRemoveBlankRows()) {
            return;
        }

        final int firstMissingRowIndex = previousRowIndex + 1;
        for (int rowIndex = firstMissingRowIndex; rowIndex < currentRowIndex; rowIndex++) {
            if (!isHiddenRow(rowIndex)) {
                stringRowConsumer.accept(null);
            }
        }
    }

    private boolean shouldSkipCell(int rowNum, int columnIndex) {
        return isHiddenRow(rowNum) || isHiddenColumn(columnIndex);
    }

    @Override
    public void cell(int rowNum, int columnIndex, String formattedValue, XSSFComment comment) {
        if (shouldSkipCell(rowNum, columnIndex)) {
            return;
        }
        super.cell(rowNum, columnIndex, formattedValue, comment);
    }

    @Override
    protected boolean shouldAcceptRow(int rowNum) {
        return !isHiddenRow(rowNum);
    }

    @Override
    protected boolean shouldAcceptColumn(int columnIndex) {
        return !isHiddenColumn(columnIndex);
    }
}