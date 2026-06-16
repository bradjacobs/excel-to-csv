/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.shared;

import org.apache.poi.ss.usermodel.DataFormatter;

public final class ExcelDataFormatters {

    // set 'true' to get actual visible value in a formula cell, and not the raw formula itself.
    private static final boolean USE_CACHED_VALUES_FOR_FORMULA_CELLS = true;
    private static final boolean EMULATE_CSV = true;

    private ExcelDataFormatters() {}

    public static DataFormatter standard() {
        return configure(new DataFormatter(EMULATE_CSV));
    }

    public static DataFormatter withDateWindowing(boolean uses1904DateWindowing) {
        return configure(new WindowingDataFormatter(uses1904DateWindowing));
    }

    private static DataFormatter configure(DataFormatter formatter) {
        formatter.setUseCachedValuesForFormulaCells(USE_CACHED_VALUES_FOR_FORMULA_CELLS);
        return formatter;
    }

    private static final class WindowingDataFormatter extends DataFormatter {

        private final boolean uses1904DateWindowing;

        private WindowingDataFormatter(boolean uses1904DateWindowing) {
            super(EMULATE_CSV);
            this.uses1904DateWindowing = uses1904DateWindowing;
        }

        @Override
        public String formatRawCellContents(double value, int formatIndex, String formatString) {
            return formatRawCellContents(value, formatIndex, formatString, uses1904DateWindowing);
        }
    }
}
