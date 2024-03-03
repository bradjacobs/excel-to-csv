/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

class ExcelSheetReader
{
    private static final boolean EMULATE_CSV = true;
    private static final DataFormatter EXCEL_DATA_FORMATTER = new DataFormatter(EMULATE_CSV);

    // "special" whitespace characters that will be converted
    //   to a "normal" whitespace character.  (*) means Character.isWhitespace() == false
    private static final Character[] SPECIAL_WHITESPACE_CHARS = {
            '\u00a0', // NON_BREAKING SPACE (*),
            '\u2002', // EN SPACE
            '\u2003', // EM SPACE
            '\u2004', // THREE-PER-EM SPACE
            '\u2005', // FOUR-PER-EM SPACE
            '\u2006', // SIX-PER-EM SPACE
            '\u2007', // FIGURE SPACE (*)
            '\u2008', // PUNCTUATION SPACE
            '\u2009', // THIN SPACE
            '\u200a', // HAIR SPACE
            '\u200b', // ZERO-WIDTH SPACE (*)
            '\u2800'  // BRAILLE SPACE (*)
    };

    // note: can fix syntax when upgrade the JDK version
    private static final Set<Character> SPECIAL_SPACE_CHAR_SET = new HashSet<>(Arrays.asList(SPECIAL_WHITESPACE_CHARS));

    private final boolean skipEmptyRows;

    protected ExcelSheetReader(boolean skipEmptyRows) {
        this.skipEmptyRows = skipEmptyRows;
    }

    /**
     * Create CSV data from the given Excel Sheet
     * @param sheet Excel Sheet
     * @return 2-D array representing CSV format
     *   each row will have the same number of columns
     */
    public String[][] convertToCsvData(Sheet sheet)  {
        List<String[]> excelListData = convertToCsvDataList(sheet);
        return excelListData.toArray(new String[0][0]);
    }

    public List<String[]> convertToCsvDataList(Sheet sheet) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet parameter cannot be null.");
        }

        // NOTE: need to add 1 to the lastRowNum to make sure you don't skip the last row
        //  (however doesn't seem to need for this when using row.getLastCellNum, which seems odd)
        int numOfRows = sheet.getLastRowNum() + 1;

        // first scan the rows to find the max column width
        int maxColumn = getMaxColumn(sheet, numOfRows);

        List<String[]> csvData = new ArrayList<>(numOfRows);

        // NOTE: avoid using "sheet.iterator" when looping through rows,
        //   b/c it can bail out early when it encounters the first empty line
        //   (even if there is more data rows remaining)
        for (int i = 0; i < numOfRows; i++) {
            String[] rowValues = new String[maxColumn];
            Row row = sheet.getRow(i);
            int columnCount = 0;
            // must check for null b/c a blank/empty row can (sometimes) return as null.
            if (row != null) {
                columnCount = Math.min(row.getLastCellNum(), maxColumn);
                for (int j = 0; j < columnCount; j++) {
                    rowValues[j] = getCellValue(row.getCell(j));
                }
            }

            // fill any 'extra' column cells with blank.
            for (int j = columnCount; j < maxColumn; j++) {
                rowValues[j] = "";
            }

            // ignore empty row if necessary
            if (this.skipEmptyRows && isEmptyRow(rowValues)) {
                continue;
            }
            csvData.add(rowValues);
        }
        return csvData;
    }

    /**
     * Iterate through the rows to find the max column width
     * @param sheet sheet
     * @param numOfRows total Number of Rows
     * @return max column
     */
    private int getMaxColumn(Sheet sheet, int numOfRows) {
        int maxColumn = 0;
        for (int i = 0; i < numOfRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                int currentRowCellCount = row.getLastCellNum();
                if (currentRowCellCount > maxColumn) {
                    //  Sometimes a row is detected with more columns, but the 'extra'
                    //    column values are actually blank.  Therefore, double check if this is
                    //    the case and adjust accordingly.
                    for (int j = currentRowCellCount - 1; j >= maxColumn; j--) {
                        String cellValue = getCellValue(row.getCell(j));
                        if (cellValue.length() > 0) {
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

    private boolean isEmptyRow(String[] rowData) {
        for (String r : rowData) {
            if (r != null && r.length() > 0) {
                return false;
            }
        }
        return true;
    }

    /**
     * Gets the string representation of the value in the cell
     *   (where the cell value is what you "physically see" in the cell)
     *  NOTE: dates & numbers should retain their original formatting.
     * @param cell excel cell
     * @return string representation of the cell.
     */
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        FormulaEvaluator evaluator = null;  // always use null for non-formula cells
        if (CellType.FORMULA.equals(cell.getCellType())) {
            evaluator = formulaEvaluator;
        }

        String cellValue = EXCEL_DATA_FORMATTER.formatCellValue(cell, evaluator);
        // if there are any special "nbsp whitespace characters", replace w/ normal whitespace
        // then return 'trimmed' value
        String sanitizedCellValue = sanitizeSpecialWhitespaceCharaters(cellValue);
        return sanitizedCellValue.trim();
    }

    /**
     * Replace any "speical/extended" whitespace characters with the
     *   basic whitespace character 0x20
     * @param input string to sanitize
     * @return string with whitespace chars replaces (if any were found)
     */
    private String sanitizeSpecialWhitespaceCharaters(String input) {
        // TODO: probably not the best way to do this!... but works for now.
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < input.length(); i++) {
            char inputCharacter = input.charAt(i);
            if (SPECIAL_SPACE_CHAR_SET.contains(inputCharacter)) {
                sb.append(' ');
            }
            else {
                sb.append(inputCharacter);
            }
        }
        return sb.toString();
    }

    /**
     * NOTE: this formulaEvaluator was copied directly from the "dummyEvaluator" in org.apache.poi.ss.util.SheetUtil,
     *  b/c it's private and the need for it is exactly the same.
     *  Namely: "...returns formula string for formula cells. Dummy evaluator makes it to format the cached formula result."
     *  @see org.apache.poi.ss.util.SheetUtil
     */
    private static final FormulaEvaluator formulaEvaluator = new FormulaEvaluator() {
        @Override public void clearAllCachedResultValues() {}
        @Override public void notifySetFormula(Cell cell) {}
        @Override public void notifyDeleteCell(Cell cell) {}
        @Override public void notifyUpdateCell(Cell cell) {}
        @Override public CellValue evaluate(Cell cell) { return null; }
        @Override public Cell evaluateInCell(Cell cell) { return null; }
        @Override public void setupReferencedWorkbooks(Map<String, FormulaEvaluator> workbooks) {}
        @Override public void setDebugEvaluationOutputForNextEval(boolean value) {}
        @Override public void setIgnoreMissingWorkbooks(boolean ignore) {}
        @Override public void evaluateAll() {}
        @Override public CellType evaluateFormulaCell(Cell cell) { return cell.getCachedFormulaResultType(); }
        @Deprecated public CellType evaluateFormulaCellEnum(Cell cell) { return evaluateFormulaCell(cell); }
    };
}
