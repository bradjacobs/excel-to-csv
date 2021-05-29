/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package bwj.util.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

class ExcelSheetReader
{
    private static final boolean EMULATE_CSV = true;
    private static final DataFormatter EXCEL_DATA_FORMATTER = new DataFormatter(EMULATE_CSV);


    private final boolean skipEmptyRows;
    private final ValueQuoter valueQuoter;


    protected ExcelSheetReader(boolean skipEmptyRows, QuoteMode quoteMode)
    {
        this.skipEmptyRows = skipEmptyRows;
        this.valueQuoter = new ValueQuoter(quoteMode);
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

    private boolean isEmptyRow(String[] rowData)
    {
        for (String r : rowData) {
            if (r != null && r.length() > 0) {
                return false;
            }
        }
        return true;
    }


    public List<String[]> convertToCsvDataList(Sheet sheet) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet parameter cannot be null.");
        }

        // NOTE: need to add 1 to the lastRowNum to make sure you don't skip the last row
        //  (however doesn't seem the need for this when using row.getLastCellNum, which seems odd)
        int numOfRows = sheet.getLastRowNum() + 1;

        // NOTE: avoid using "sheet.iterator" when looping thru rows,
        //   b/c it can bail out early when it encounters the first empty line
        //   (even if there is more data rows remaining)
        int maxColumn = 0;

        // first iterate thru the rows to find the max column width
        for (int i = 0; i < numOfRows; i++)
        {
            Row row = sheet.getRow(i);
            if (row != null)
            {
                maxColumn = Math.max(maxColumn, row.getLastCellNum());
            }
        }

        List<String[]> csvData = new ArrayList<>(numOfRows);

        for (int i = 0; i < numOfRows; i++)
        {
            int columnCount = 0;
            String[] rowValues = new String[maxColumn];

            Row row = sheet.getRow(i);
            // must check for null b/c a blank/empty row can return as null.
            if (row != null)
            {
                columnCount = row.getLastCellNum();
                for (int j = 0; j < columnCount; j++)
                {
                    String cellValue = getCellValue(row.getCell(j));
                    // wrap the cellValue inside quotes IFF if configured and necessary
                    cellValue = valueQuoter.applyCsvQuoting(cellValue);
                    rowValues[j] = cellValue;
                }
            }

            if (columnCount < maxColumn)
            {
                for (int j = columnCount; j < maxColumn; j++) {
                    rowValues[j] = "";
                }
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
     * Gets the string representation of the value in the cell (where "value" is what you physically see in the cell)
     *  NOTE; dates & numbers should retain their original formatting.
     * @param cell excel cell
     * @return string representation of the cell.
     */
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        String formattedCellValue;
        if (cell.getCellType() != null && cell.getCellType().equals(CellType.FORMULA)) {
            formattedCellValue = EXCEL_DATA_FORMATTER.formatCellValue(cell, formulaEvaluator);
        }
        else {
            formattedCellValue = EXCEL_DATA_FORMATTER.formatCellValue(cell);
        }

        // NOTE: maybe not 'officially' accurate, but going to assume that extra leading/trailing space should not be there.
        return formattedCellValue.trim();
    }


    /**
     * NOTE: this formulaEvaluator was copied directly from the "dummyEvaluator" in org.apache.poi.ss.util.SheetUtil,
     *  b/c it's private and the need for it is exactly the same.
     *  Namely "...returns formula string for formula cells. Dummy evaluator makes it to format the cached formula result."
     *  @see org.apache.poi.ss.util.SheetUtil
     */
    private static final FormulaEvaluator formulaEvaluator = new FormulaEvaluator() {
        @Override
        public void clearAllCachedResultValues() {}
        @Override
        public void notifySetFormula(Cell cell) {}
        @Override
        public void notifyDeleteCell(Cell cell) {}
        @Override
        public void notifyUpdateCell(Cell cell) {}
        @Override
        public CellValue evaluate(Cell cell) { return null; }
        @Override
        public Cell evaluateInCell(Cell cell) { return null; }
        @Override
        public void setupReferencedWorkbooks(Map<String, FormulaEvaluator> workbooks) {}
        @Override
        public void setDebugEvaluationOutputForNextEval(boolean value) {}
        @Override
        public void setIgnoreMissingWorkbooks(boolean ignore) {}
        @Override
        public void evaluateAll() {}
        @Override
        public CellType evaluateFormulaCell(Cell cell) { return cell.getCachedFormulaResultType(); }
        @Deprecated
        @Override
        public CellType evaluateFormulaCellEnum(Cell cell) { return evaluateFormulaCell(cell); }
    };
}
