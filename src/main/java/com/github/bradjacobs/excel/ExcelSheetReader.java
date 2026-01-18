/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.function.Consumer;
import java.util.function.IntPredicate;
import java.util.function.Predicate;
import java.util.function.UnaryOperator;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.DASHES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.QUOTES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.SPACES;

// TODO javadocs
public class ExcelSheetReader {
    protected final boolean removeBlankRows;
    protected final boolean removeBlankColumns;
    protected final boolean removeInvisibleCells;
    protected final CellValueReader cellValueReader;

    private ExcelSheetReader(Builder builder) {
        this.removeBlankRows = builder.removeBlankRows;
        this.removeBlankColumns = builder.removeBlankColumns;
        this.removeInvisibleCells = builder.removeInvisibleCells;
        this.cellValueReader = new CellValueReader(builder.autoTrim, builder.charSanitizeFlags);
    }

    /**
     * Create 2-D data matrix from the given Excel Sheet
     * @param sheet Excel Sheet
     * @return 2-D array representing CSV format
     * each row will have the same number of columns
     */
    public String[][] convertToDataMatrix(Sheet sheet) {
        List<String[]> excelListData = convertToDataMatrixList(sheet);
        return excelListData.toArray(new String[0][0]);
    }

    public List<String[]> convertToDataMatrixList(Sheet sheet) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet parameter cannot be null.");
        }

        // grab all the rows from the sheet
        List<Row> rowList = getRows(sheet);
        // first scan the rows to find the max column width
        int maxColumn = getMaxColumn(rowList);
        // get all the column (indexes) that are to be read
        int[] availableColumns = getAvailableColumns(sheet, maxColumn);

        return convertToDataMatrixList(rowList, availableColumns);
    }

    protected List<String[]> convertToDataMatrixList(List<Row> rowList, int[] availableColumns) {
        // if there are no available columns then bail early.
        if (availableColumns.length == 0) {
            return Collections.emptyList();
        }

        int totalColumnCount = availableColumns.length;
        int lastColumnIndex = availableColumns[totalColumnCount-1];

        StringArrayRowConsumer stringArrayRowConsumer =
                new StringArrayRowConsumer(this.removeBlankRows, this.removeBlankColumns);
        // NOTE: avoid using "sheet.iterator" when looping through rows,
        //   because it can bail out early when it encounters the first empty line
        //   (even if there are more data rows remaining)
        for (Row row : rowList) {
            String[] rowValues = new String[totalColumnCount];

            // must check for null because a blank/empty row can (sometimes) be null.
            int columnIndex = 0;
            if (row != null) {
                int rowColumnCount = Math.max(row.getLastCellNum(), 0);
                int lastRowColumnIndex = Math.min(rowColumnCount-1, lastColumnIndex);

                for (int availableColumn : availableColumns) {
                    if (availableColumn > lastRowColumnIndex) {
                        break;
                    }
                    rowValues[columnIndex++] = getCellValue(row.getCell(availableColumn));
                }
            }

            while (columnIndex < totalColumnCount) {
                rowValues[columnIndex++] = "";
            }
            stringArrayRowConsumer.accept(rowValues);
        }

        // call finalizeRows to handle any extra remove blank row/column logic
        stringArrayRowConsumer.finalizeRows();

        // lastly return the all rows
        return stringArrayRowConsumer.getRows();
    }

    /**
     * Iterate through the rows to find the max column that contains a value
     * @param rowList list of rows for a sheet
     * @return max column
     */
    protected int getMaxColumn(List<Row> rowList) {
        int maxColumn = 0;
        for (Row row : rowList) {
            if (row != null) {
                int currentRowCellCount = row.getLastCellNum();
                if (currentRowCellCount > maxColumn) {
                    //  Sometimes a row is detected with more columns, but the 'extra'
                    //    column values are actually blank.  Therefore, double check if this is
                    //    the case and adjust accordingly.
                    for (int j = currentRowCellCount - 1; j >= maxColumn; j--) {
                        String cellValue = getCellValue(row.getCell(j));
                        if (!cellValue.isEmpty()) {
                            break;
                        }
                        currentRowCellCount--;
                    }
                    maxColumn = Math.max(maxColumn, currentRowCellCount);
                }
            }
        }
        return maxColumn;
    }

    /**
     * Gets the string representation of the value in the cell
     *   (where the cell value is what you "physically see" in the cell)
     * NOTE: dates & numbers should retain their original formatting.
     * @param cell excel cell
     * @return string representation of the cell.
     */
    protected String getCellValue(Cell cell) {
        return cellValueReader.getCellValue(cell);
    }

    /**
     * Method to grab all the rows for the sheet ahead of time
     *   NOTE: some elements in the result list could be 'null'
     *   (nulls are usually 'default unaltered rows')
     * @param sheet input Excel Sheet
     * @return list of rows
     */
    protected List<Row> getRows(Sheet sheet) {
        // NOTE: need to add 1 to the lastRowNum to make sure don't skip the last row
        //  (however doesn't seem to need for this when using row.getLastCellNum, which seems odd)
        int numOfRows = sheet.getLastRowNum() + 1;

        Predicate<Row> rowFilterPredicate = removeInvisibleCells ?
                r -> r == null || !r.getZeroHeight() :
                r -> true;
        return IntStream.range(0, numOfRows)
                .mapToObj(sheet::getRow)
                .filter(rowFilterPredicate)
                .collect(Collectors.toList());
    }

    /**
     * Determine which columns of the sheet should be read
     * @param sheet sheet
     * @param maxColumn max column count
     * @return int array of column indices to be read
     */
    protected int[] getAvailableColumns(Sheet sheet, int maxColumn) {
        IntPredicate intPredicate = removeInvisibleCells ?
                idx -> !sheet.isColumnHidden(idx) :
                idx -> true;
        return IntStream.range(0, maxColumn).filter(intPredicate).toArray();
    }

    /**
     * Consumer class to add the String[] rows, which will handle
     * any additional logic regarding the removal of blank rows or columns (as needed)
     */
    private static class StringArrayRowConsumer implements Consumer<String[]> {
        private final List<String[]> rowList = new ArrayList<>();
        private final boolean removeBlankRows; // flag to remove blank rows
        private final boolean removeBlankColumns; // flag to remove blank columns
        private boolean[] columnsHaveDataArray = null; // tracks which columns have a value
        private int numberOfColumnsHaveData = 0; // tracks total number of columns that have non-blank value.

        public StringArrayRowConsumer(boolean removeBlankRows, boolean removeBlankColumns) {
            this.removeBlankRows = removeBlankRows;
            this.removeBlankColumns = removeBlankColumns;
        }

        /**
         * Consumes all the String[] rows read from the Excel file
         * _ASSERT_ that all rows passed in will have the same length.
         * @param row String[] row
         */
        @Override
        public void accept(String[] row) {
            // Remove blank rows if configured
            if (this.removeBlankRows && isEmptyRow(row))
                return;

            // if removing blank columns, then keep track which columns contain values.
            if (this.removeBlankColumns) {
                if (columnsHaveDataArray == null) {
                    columnsHaveDataArray = new boolean[row.length];
                }
                if (numberOfColumnsHaveData < columnsHaveDataArray.length) {
                    for (int i = 0; i < columnsHaveDataArray.length; i++) {
                        if (!columnsHaveDataArray[i] && !row[i].isEmpty()) {
                            columnsHaveDataArray[i] = true;
                            numberOfColumnsHaveData++;
                        }
                    }
                }
            }
            this.rowList.add(row);
        }

        public List<String[]> getRows() {
            return Collections.unmodifiableList(rowList);
        }

        /**
         * To be called after all rows have been added for any post-processing (if necessary)
         */
        public void finalizeRows() {
            // remove any trailing blank rows (regardless of 'removeBlankRows' value)
            while (!rowList.isEmpty() && isEmptyRow(rowList.get(rowList.size()-1))) {
                rowList.remove(rowList.size()-1);
            }

            // remove any blank columns (if necessary)
            if (removeBlankColumns &&
                    !rowList.isEmpty() &&
                    columnsHaveDataArray != null &&
                    numberOfColumnsHaveData < columnsHaveDataArray.length) {
                rowList.replaceAll(new FilterColumnsOperator(columnsHaveDataArray, numberOfColumnsHaveData));
            }
        }
    }

    // Helper class to remove blank column(s) from each String[] row
    private static class FilterColumnsOperator implements UnaryOperator<String[]> {
        private final boolean[] columnHasData;
        private final int columnsWithDataCount;

        public FilterColumnsOperator(boolean[] columnHasData, int columnsWithDataCount) {
            this.columnHasData = columnHasData;
            this.columnsWithDataCount = columnsWithDataCount;
        }

        @Override
        public String[] apply(String[] row) {
            String[] filtered = new String[columnsWithDataCount];
            for (int i = 0, idx = 0; i < columnHasData.length && i < row.length; i++) {
                if (columnHasData[i]) {
                    filtered[idx++] = row[i];
                }
            }
            return filtered;
        }
    }

    /**
     * Check if row is considered to be 'blank/empty'
     * @param rowData data row
     * @return true if 'empty'
     */
    protected static boolean isEmptyRow(String[] rowData) {
        for (String r : rowData) {
            if (r != null && !r.isEmpty()) {
                return false;
            }
        }
        return true;
    }

    public static Builder builder() {
        return new Builder();
    }

    protected static abstract class AbstractSheetConfigBuilder<T extends AbstractSheetConfigBuilder<T>> {
        protected boolean autoTrim = true;
        protected boolean removeBlankRows = false;
        protected boolean removeBlankColumns = false;
        protected boolean removeInvisibleCells = false;
        protected Set<CharSanitizeFlag> charSanitizeFlags
                = new HashSet<>(SpecialCharacterSanitizer.DEFAULT_FLAGS);

        protected abstract T self();

        /**
         * Whether to trim whitespace on cell values
         * @param autoTrim (defaults to true)
         */
        public T autoTrim(boolean autoTrim) {
            this.autoTrim = autoTrim;
            return self();
        }

        /**
         * Whether to remove any blank rows.
         * @param removeBlankRows (defaults to false)
         */
        public T removeBlankRows(boolean removeBlankRows) {
            this.removeBlankRows = removeBlankRows;
            return self();
        }

        /**
         * Whether to remove any blank columns.
         * @param removeBlankColumns (defaults to false)
         */
        public T removeBlankColumns(boolean removeBlankColumns) {
            this.removeBlankColumns = removeBlankColumns;
            return self();
        }

        public T removeInvisibleCells(boolean removeInvisibleCells) {
            this.removeInvisibleCells = removeInvisibleCells;
            return self();
        }

        public T sanitizeSpaces(boolean sanitizeSpaces) {
            return setSanitizeFlag(SPACES, sanitizeSpaces);
        }

        public T sanitizeQuotes(boolean sanitizeQuotes) {
            return setSanitizeFlag(QUOTES, sanitizeQuotes);
        }

        public T sanitizeDiacritics(boolean sanitizeDiacritics) {
            return setSanitizeFlag(BASIC_DIACRITICS, sanitizeDiacritics);
        }

        public T sanitizeDashes(boolean sanitizeDashes) {
            return setSanitizeFlag(DASHES, sanitizeDashes);
        }

        private T setSanitizeFlag(SpecialCharacterSanitizer.CharSanitizeFlag flag, boolean shouldAdd) {
            if (shouldAdd) {
                this.charSanitizeFlags.add(flag);
            }
            else {
                this.charSanitizeFlags.remove(flag);
            }
            return self();
        }

        // setting the entire Set of SpecialCharSanitizers has limited access.
        protected T charSanitizeFlags(Set<SpecialCharacterSanitizer.CharSanitizeFlag> charSanitizeFlags) {
            this.charSanitizeFlags =charSanitizeFlags;
            return self();
        }
    }

    public static class Builder extends AbstractSheetConfigBuilder<Builder> {
        @Override
        protected Builder self() {
            return this;
        }

        public ExcelSheetReader build() {
            return new ExcelSheetReader(this);
        }
    }
}
