package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.CellValueReader;
import com.github.bradjacobs.excel.SheetConfig;
import com.github.bradjacobs.excel.StringRowConsumer;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.ArrayList;
import java.util.List;

// todo: some of the logic is a little squirrely,
//   but want to first confirm all edge cases work correctly
//   and unittests are in place before doing refactor.
class SpecialSheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    protected static final String EMPTY_CELL = "";

    private static final String EXCEL_ERROR_PREFIX = "ERROR:#";
    private static final int EXCEL_ERROR_PREFIX_STRIP_LENGTH = 6;

    protected final SheetConfig sheetConfig;
    protected final CellValueReader cellValueReader;
    protected final StringRowConsumer stringRowConsumer;

    protected final List<String> currentRowValues = new ArrayList<>();

    public SpecialSheetHandler(SheetConfig sheetConfig, StringRowConsumer stringRowConsumer) {
        this.sheetConfig = sheetConfig;
        this.cellValueReader = new CellValueReader(
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
        if (StringUtils.isEmpty(cellReference)) {
            // todo: just throw an exception if missing cellRefernce
            //    other online 'solutions' show a way to manully track the column index,
            //    but it seems to often be incorrect and would write cell data in incorret column.
            throw new IllegalStateException(
                    "Unable to parse Excel Sheet. " +
                            "A cell value was encountered without a cellReference.  " +
                            "See 'Known Issues' for more details.");
        }
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
        return removeExcelErrorPrefix(cellValueReader.sanitizeCellValue(cellValue));
    }

    /**
     * remove the first part of error string to be consistent
     * with the behavior of reading cell values from Cell/Row/Sheet objects
     */
    private String removeExcelErrorPrefix(String input) {
        // note: checking the first 7 characters, but only removing the first 6.
        if (input.startsWith(EXCEL_ERROR_PREFIX)) {
            return input.substring(EXCEL_ERROR_PREFIX_STRIP_LENGTH);
        }
        return input;
    }

    private void clearCurrentRow() {
        currentRowValues.clear();
    }

    @Override
    public void endSheet() {
        // nothing to do
    }
}
