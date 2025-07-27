/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.util.Set;

class CellValueReader {
    private static final boolean EMULATE_CSV = true;
    private static final DataFormatter EXCEL_DATA_FORMATTER = new DataFormatter(EMULATE_CSV);
    static {
        // set true to get actual visible value in a formula cell, and not the raw formula itself.
        EXCEL_DATA_FORMATTER.setUseCachedValuesForFormulaCells(true);
    }

    private final boolean autoTrim;
    private final SpecialCharacterSanitizer specialCharSanitizer;

    public CellValueReader(boolean autoTrim, Set<CharSanitizeFlag> charSanitizeFlags) {
        this.autoTrim = autoTrim;
        this.specialCharSanitizer = new SpecialCharacterSanitizer(charSanitizeFlags);
    }

    /**
     * Gets the string representation of the value in the cell
     *   (where the cell value is what you "physically see" in the cell)
     *  NOTE: dates & numbers should retain their original formatting.
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

        // if there are any certain special unicode characters (like nbsp or smart quotes),
        // replace w/ normal character equivalent
        cellValue = specialCharSanitizer.sanitize(cellValue);

        if (this.autoTrim) {
            cellValue = cellValue.trim();
        }
        return cellValue;
    }
}
