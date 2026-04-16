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
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.Validate;

import java.io.IOException;
import java.net.URL;
import java.nio.file.Path;
import java.util.List;

/**
 * ExcelProcessor
 */
public class ExcelProcessor implements ExcelSheetReader {
    private static final String XLSX_EXTENSION = "xlsx";

    private final boolean useAdvancedReader;
    private final ExcelSheetReader standardExcelSheetReader;
    private final ExcelSheetReader advancedExcelSheetReader;

    private ExcelProcessor(Builder builder) {
        this.useAdvancedReader = builder.useAdvancedReader;
        SheetConfig sheetConfig = builder.buildConfig();
        this.standardExcelSheetReader = new StandardExcelSheetReader(sheetConfig);
        this.advancedExcelSheetReader = new AdvancedExcelSheetReader(sheetConfig);
    }

    @Override
    public List<SheetContent> readSheets(ExcelSheetReadRequest request) throws IOException {
        Validate.isTrue(request != null, "Request cannot be null");
        return getSheetReaderForRequest(request).readSheets(request);
    }

    private ExcelSheetReader getSheetReaderForRequest(ExcelSheetReadRequest request) {
        return shouldUseAdvancedReader(request)
                ? this.advancedExcelSheetReader
                : this.standardExcelSheetReader;
    }

    private boolean shouldUseAdvancedReader(ExcelSheetReadRequest request) {
        return this.useAdvancedReader && hasXlsxExtension(request);
    }

    private boolean hasXlsxExtension(ExcelSheetReadRequest request) {
        String sourceLocation = resolveSourceLocation(request);
        return XLSX_EXTENSION.equalsIgnoreCase(FilenameUtils.getExtension(sourceLocation));
    }

    private String resolveSourceLocation(ExcelSheetReadRequest request) {
        Path path = request.getPath();
        if (path != null) {
            return path.toString();
        }
        URL url = request.getUrl();
        if (url != null) {
            return url.toString();
        }
        return "";
    }

    public static Builder builder() {
        return new Builder();
    }

    // this builder extends the abstract class to allow any of the
    //   AbstractSheetConfigBuilder values to be set on this Builder as well.
    public static class Builder extends AbstractSheetConfigBuilder<ExcelProcessor, Builder> {
        private boolean useAdvancedReader = true;

        @Override
        protected Builder self() {
            return this;
        }

        /**
         * Toggle using the advanced reader
         * (typically only used for testing convenience)
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
