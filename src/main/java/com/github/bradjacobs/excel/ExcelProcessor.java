/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.advanced.AdvancedExcelSheetReader;
import com.github.bradjacobs.excel.api.ExcelSheetReader;
import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.AbstractExcelSheetReader.AbstractSheetConfigBuilder;
import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import com.github.bradjacobs.excel.standard.StandardExcelSheetReader;
import org.apache.commons.lang3.Validate;

import java.io.IOException;
import java.util.List;

/**
 * ExcelProcessor
 */
public class ExcelProcessor implements ExcelSheetReader {

    private static final String XLSX_EXTENSION = ".xlsx";

    private final boolean useAdvancedReader;
    private final ExcelSheetReader standardExcelSheetReader;
    private final ExcelSheetReader advancedExcelSheetReader;

    private ExcelProcessor(Builder builder) {
        this.useAdvancedReader = builder.useAdvancedReader;
        SheetConfig sheetConfig = builder.buildConfig();
        this.standardExcelSheetReader = new StandardExcelSheetReader(sheetConfig);
        this.advancedExcelSheetReader = new AdvancedExcelSheetReader(sheetConfig);
    }

    public List<SheetContent> readSheets(ExcelSheetReadRequest request) throws IOException {
        Validate.isTrue(request != null, "Request cannot be null");
        ExcelSheetReader sheetReader = selectSheetReader(request);
        return sheetReader.readSheets(request);
    }

    private ExcelSheetReader selectSheetReader(ExcelSheetReadRequest request) {
        if (shouldUseAdvancedReader(request)) {
            return this.advancedExcelSheetReader;
        }
        return this.standardExcelSheetReader;
    }

    private boolean shouldUseAdvancedReader(ExcelSheetReadRequest request) {
        // todo - it's possible that password-protected can use
        //   the advanced reader, but those checks ore not presently in place.
        if (!this.useAdvancedReader || request.getPassword() != null) {
            return false;
        }
        return isXlsxSource(request);
    }

    private boolean isXlsxSource(ExcelSheetReadRequest request) {
        if (request.getPath() != null) {
            return request.getPath().toString().toLowerCase().endsWith(XLSX_EXTENSION);
        }
        return request.getUrl() != null && request.getUrl().toString().toLowerCase().endsWith(XLSX_EXTENSION);
    }

    public static Builder builder() {
        return new Builder();
    }

    // this builder extends the abstract class to allow any of the
    //   AbstractSheetConfigBuilder values to be set on this Builder as well.
    public static class Builder extends AbstractSheetConfigBuilder<ExcelProcessor, Builder> {
        private boolean useAdvancedReader = true; // use the advanced reader (if possible)

        @Override
        protected Builder self() {
            return this;
        }

        /**
         * Toggle using the advanced reader
         *   (typically only used for testing convenience)
         */
        public Builder useAdvancedReader(boolean useAdvancedReader) {
            this.useAdvancedReader = useAdvancedReader;
            return this;
        }

        @Override
        public ExcelProcessor build() {
            return new ExcelProcessor(this);
        }
    }
}
