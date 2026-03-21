/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.advanced.AdvancedExcelSheetReader;
import com.github.bradjacobs.excel.api.ExcelSheetReader;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.AbstractExcelSheetReader;
import com.github.bradjacobs.excel.standard.StandardExcelSheetReader;

// NOTE: maybe in future may change to enforce
//   the usage of the builder, but right now it's optional
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

    public static Builder standard() {
        return new Builder().type(ReaderType.STANDARD);
    }

    public static Builder advanced() {
        return new Builder().type(ReaderType.ADVANCED);
    }

    public static Builder builder() {
        return standard(); // standard is the 'default'
    }

    public static class Builder extends AbstractExcelSheetReader.AbstractSheetConfigBuilder<ExcelSheetReader, Builder> {
        private ReaderType type = ReaderType.STANDARD;

        public Builder type(ReaderType type) {
            if (type == null) {
                throw new IllegalArgumentException("Reader type cannot be null");
            }
            this.type = type;
            return this;
        }

        @Override
        protected Builder self() {
            return this;
        }

        @Override
        public ExcelSheetReader build() {
            return ExcelSheetReaderFactory.create(type, buildConfig());
        }
    }
}