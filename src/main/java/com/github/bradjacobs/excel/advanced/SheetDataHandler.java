/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.CellValueSanitizer;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.ArrayList;
import java.util.List;

// todo: some of the logic is a little squirrelly,
//   but want to first confirm all edge cases work correctly
//   and unittests are in place before doing refactor.
class SheetDataHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    protected static final String EMPTY_CELL = "";
    private static final String EXCEL_ERROR_PREFIX = "ERROR:";
    private static final String MISSING_CELL_REF_MSG = "Unable to parse Excel Sheet. " +
            "A cell value was encountered without a cellReference.  " +
            "See 'Known Issues' for more details.";

    protected final SheetConfig sheetConfig;
    protected final CellValueSanitizer cellValueSanitizer;
    protected final StringRowConsumer stringRowConsumer;

    protected final List<String> currentRowValues = new ArrayList<>();

    public SheetDataHandler(SheetConfig sheetConfig, StringRowConsumer stringRowConsumer) {
        this.sheetConfig = sheetConfig;
        this.cellValueSanitizer = new CellValueSanitizer(
                sheetConfig.isAutoTrim(),
                sheetConfig.getCharSanitizeFlags()
        );
        this.stringRowConsumer = stringRowConsumer;
    }

    public String[][] getMatrix() {
        return stringRowConsumer.generateMatrix();
    }

    @Override
    public void startRow(int rowNum) {
        fillMissingRows(rowNum);
    }

    @Override
    public void endRow(int rowNum) {
        if (shouldAcceptRow(rowNum)) {
            stringRowConsumer.accept(currentRowValues);
        }
        clearCurrentRow();
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        requireCellReference(cellReference);
        CellAddress cellAddress = new CellAddress(cellReference);
        cell(cellAddress.getRow(), cellAddress.getColumn(), formattedValue, comment);
    }

    protected void cell(int rowNum, int columnIndex, String formattedValue, XSSFComment comment) {
        fillMissingColumnsUpTo(columnIndex);
        currentRowValues.add(sanitizeCellValue(formattedValue));
    }

    protected void fillMissingRows(int rowNum) {
        if (sheetConfig.isRemoveBlankRows()) {
            return;
        }

        // add any filler blank row (if necessary)
        while (stringRowConsumer.getRowCount() < rowNum) {
            stringRowConsumer.accept(null);
        }
    }

    protected boolean shouldAcceptRow(int rowNum) {
        return true;
    }

    protected boolean shouldAcceptColumn(int columnIndex) {
        return true;
    }

    protected void fillMissingColumnsUpTo(int columnIndex) {
        // fill in any blanks between values in a row (if necessary)
        for (int col = currentRowValues.size(); col < columnIndex; col++) {
            if (shouldAcceptColumn(col)) {
                currentRowValues.add(EMPTY_CELL);
            }
        }
    }

    protected String sanitizeCellValue(String cellValue) {
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

    /**
     * Check existence of cellReference.
     * Fail if it doesn't exist because the cell position
     * cannot be recovered reliably.
     * @param cellReference cell reference
     */
    private void requireCellReference(String cellReference) {
        if (StringUtils.isEmpty(cellReference)) {
            throw new IllegalStateException(MISSING_CELL_REF_MSG);
        }
    }
}
