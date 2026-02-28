/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * A more bare-bones example of using the POI library
 * for reading sheet data from an Excel file.
 *
 * NOTE_1:
 *   Generally it's best to avoid row/column iterators because it can produce an
 *   unexpected result with null cells or rows.
 *
 *   Example: If an Excel sheet was like this:
 *           A     B     C    D
 *       1  dog              cat
 *       2        cow        pig
 *   Then one would _assume_ that the result would look like this:
 *       [["dog", "", "", "cat"], ["", "cow", "", "pig"]]
 *   But using iterator can result in something like this:
 *       [["dog", "cat"], ["cow", "pig"]]
 *
 *   Do note that without the iterator it's possible to have trailing blank rows
 *     after all the data is read.
 *
 * NOTE_2:
 *   set 'formatter.setUseCachedValuesForFormulaCells(true)'
 *   to get the actual value one sees in a cell (and not the raw formula)
 *
 * NOTE_3:
 *   You may (or may not) want to allow rows with different lengths.
 *   There sre 2 examples below (one for each scenario)
 *
 * NOTE_4:
 *   Be aware that it's possible for 'row.getLastCellNum()'
 *     to return -1  (though it doesn't matter for this example)
 *
 * NOTE_5:
 *   To get the correct rowCount is actually
 *     sheet.getLastRowNum()  +  1
 *   but it's _not_ correct add +1 to 'row.getLastCellNum()' (oddly enough).
 *   In the code below used 'IntStream.rangeClosed' instead of 'IntStream.range'
 *   to avoid the extra + 1
 *
 * NOTE_6:
 *   The examples below till call 'trim()' on each of the cell values.
 *   Easily removable if not desired.
 *
 * NOTE_7:
 *   'formatter.formatCellValue(null)' - returns an empty string
 */
public class SimplePoiExampleExcelReader {

    private static final boolean EMULATE_CSV = true;
    private static final DataFormatter formatter = new DataFormatter(EMULATE_CSV);
    static {
        // set to get actual cell value (and not formula)
        formatter.setUseCachedValuesForFormulaCells(true);
    }

    /**
     * given an Excel file path, return all the data from the first sheet
     * @param excelFile excel file
     * @return sheet data from the first sheet
     * @throws IOException exception
     */
    public List<List<String>> readBasicSheet(Path excelFile) throws IOException {
        DataFormatter formatter = new DataFormatter(true);
        formatter.setUseCachedValuesForFormulaCells(true);

        try (Workbook workbook = WorkbookFactory.create(excelFile.toFile())) {
            Sheet sheet = workbook.getSheetAt(0); // read first sheet
            return IntStream.rangeClosed(0, sheet.getLastRowNum())
                    .mapToObj(rowIndex -> {
                        Row row = sheet.getRow(rowIndex);
                        if (row == null) {
                            return List.<String>of();
                        }
                        return IntStream.range(0, row.getLastCellNum())
                                .mapToObj(colIndex ->
                                        formatter.formatCellValue(row.getCell(colIndex)).trim())
                                .collect(Collectors.toList());
                    })
                    .collect(Collectors.toList());
        }
    }

    /**
     * Alternate example of reading a sheet.
     * - All rows in the returned list will be the same length.
     * - Use a separate method for getCellValue
     * - Use private static formatter instance
     */
    public List<List<String>> readBasicSheetResultRowsEqualLength(Path excelFile) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(excelFile.toFile())) {
            Sheet sheet = workbook.getSheetAt(0); // read first sheet
            int maxColumns = IntStream.rangeClosed(0, sheet.getLastRowNum())
                    .mapToObj(sheet::getRow)
                    .filter(Objects::nonNull)
                    .mapToInt(Row::getLastCellNum)
                    .max()
                    .orElse(0);

            return IntStream.rangeClosed(0, sheet.getLastRowNum())
                    .mapToObj(rowIndex -> {
                        Row row = sheet.getRow(rowIndex);
                        return IntStream.range(0, maxColumns)
                                .mapToObj(colIndex -> getCellValue(row != null ? row.getCell(colIndex) : null))
                                .collect(Collectors.toList());
                    })
                    .collect(Collectors.toList());
        }
    }

    private String getCellValue(Cell cell) {
        return cell != null ? formatter.formatCellValue(cell).trim() : "";
    }
}