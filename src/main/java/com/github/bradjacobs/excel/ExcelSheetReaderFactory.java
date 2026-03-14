/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.advanced.AdvancedExcelSheetReader;
import com.github.bradjacobs.excel.api.ExcelSheetReader;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.standard.StandardExcelSheetReader;

public class ExcelSheetReaderFactory {

    public enum ReaderType {
        STANDARD,
        ADVANCED
    }

    private ExcelSheetReaderFactory() {}

    public static ExcelSheetReader create(ReaderType type, SheetConfig sheetConfig) {
        switch (type) {
            case STANDARD:
                return new StandardExcelSheetReader(sheetConfig);
            case ADVANCED:
                return new AdvancedExcelSheetReader(sheetConfig);
            default:
                throw new IllegalArgumentException("Unsupported reader type: " + type);
        }
    }
}