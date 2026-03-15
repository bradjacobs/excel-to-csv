/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced.datewindowing;

import org.apache.poi.ss.usermodel.DataFormatter;

/**
 * Special override of DataFormatter to force the value
 * of use1904Windowing and not rely on the default value.
 */
public class DateWindowingDataFormatter extends DataFormatter {

    private static final boolean EMULATE_CSV = true;

    private final boolean uses1904DateWindowing;

    public DateWindowingDataFormatter(boolean uses1904DateWindowing) {
        super(EMULATE_CSV);
        this.setUseCachedValuesForFormulaCells(true);
        this.uses1904DateWindowing = uses1904DateWindowing;
    }

    @Override
    public String formatRawCellContents(double value, int formatIndex, String formatString) {
        return formatRawCellContents(value, formatIndex, formatString, uses1904DateWindowing);
    }
}
