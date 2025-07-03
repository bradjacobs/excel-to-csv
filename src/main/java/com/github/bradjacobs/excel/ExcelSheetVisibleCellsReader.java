/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * ExcelSheetReader that only handleS 'visible' cells.
 *   i.e. the cells must have width > 0 and height > 0
 */
class ExcelSheetVisibleCellsReader extends ExcelSheetReader {

    public ExcelSheetVisibleCellsReader(
            boolean autoTrim,
            boolean skipEmptyRows,
            Set<SpecialCharacterSanitizer.CharSanitizeFlag> charSanitizeFlags) {
        super(autoTrim, skipEmptyRows, charSanitizeFlags);
    }

    /** @inheritDoc */
    @Override
    protected List<Row> getRows(Sheet sheet) {
        List<Row> originalRowList = super.getRows(sheet);

        // filter out any rows that have a 'zero height'
        return originalRowList.stream()
                .filter(r -> r == null || !r.getZeroHeight())
                .collect(Collectors.toList());
    }

    /** @inheritDoc */
    @Override
    protected int[] getAvailableColumns(Sheet sheet, int maxColumn) {
        return IntStream.range(0, maxColumn)
                .filter(idx -> !sheet.isColumnHidden(idx))
                .toArray();
    }
}
