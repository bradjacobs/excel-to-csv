/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import com.github.bradjacobs.excel.config.SanitizeType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.util.Set;

// todo: the extends class is an experiment
//   to see first hand how ok or bad the implementation turns out.
public class CellValueReader extends CellValueSanitizer {
    private static final boolean EMULATE_CSV = true;
    private static final DataFormatter EXCEL_DATA_FORMATTER = new DataFormatter(EMULATE_CSV);
    static {
        // set true to get actual visible value in a formula cell, and not the raw formula itself.
        EXCEL_DATA_FORMATTER.setUseCachedValuesForFormulaCells(true);
    }

    public CellValueReader(boolean autoTrim, Set<SanitizeType> sanitizeTypes) {
        super(autoTrim, sanitizeTypes);
    }

    /**
     * Gets the string representation of the value in the cell
     *   (where the cell value is what you "physically see" in the cell)
     * NOTE: dates and numbers should retain their original formatting.
     * @param cell excel cell
     * @return string representation of the cell.
     */
    public String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        String cellValue = EXCEL_DATA_FORMATTER.formatCellValue(cell);

        // NOTE: below is example of how to ensure 'blank value' if a custom format
        //   of '3 semicolons' ;;; was used to hide cell contents.
        //   However, it's commented out b/c negative performance impact could outweigh benefit.
        //if (cell.getCellStyle().getDataFormatString().equals(";;;")) {
        //    cellValue = "";
        //}

        // return a sanitized version of the cell value which is trimmed (if configured)
        // plus convert any special Unicode characters (like nbsp or smart quotes),
        // as necessary.
        return sanitizeCellValue(cellValue);
    }
}
