/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.api.ExcelWorkbookReader;
import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.AbstractExcelReader.AbstractSheetConfigBuilder;
import com.github.bradjacobs.excel.engine.eventmodel.AdvancedExcelReader;
import com.github.bradjacobs.excel.engine.objectmodel.StandardExcelReader;
import com.github.bradjacobs.excel.request.ExcelReadRequest;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.Validate;

import java.io.IOException;
import java.net.URL;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;

/**
 * ExcelProcessor
 */
public class ExcelProcessor implements ExcelWorkbookReader {
    private static final List<String> ADVANCED_SUPPORTED_EXTENSIONS
            = List.of("xlsx", "xlsb", "xlsm", "xltx", "xltm");

    private final boolean useAdvancedReader;
    private final ExcelWorkbookReader standardExcelWorkbookReader;
    private final ExcelWorkbookReader advancedExcelWorkbookReader;

    private ExcelProcessor(Builder builder) {
        this.useAdvancedReader = builder.useAdvancedReader;
        SheetConfig sheetConfig = builder.buildConfig();
        this.standardExcelWorkbookReader = new StandardExcelReader(sheetConfig);
        this.advancedExcelWorkbookReader = new AdvancedExcelReader(sheetConfig);
    }

    @Override
    public List<SheetContent> readSheets(ExcelReadRequest request) throws IOException {
        Validate.isTrue(request != null, "Request cannot be null");
        return getSheetReaderForRequest(request).readSheets(request);
    }

    private ExcelWorkbookReader getSheetReaderForRequest(ExcelReadRequest request) {
        return shouldUseAdvancedReader(request)
                ? this.advancedExcelWorkbookReader
                : this.standardExcelWorkbookReader;
    }

    private boolean shouldUseAdvancedReader(ExcelReadRequest request) {
        return this.useAdvancedReader && hasAdvancedSupportedExtension(request);
    }

    private boolean hasAdvancedSupportedExtension(ExcelReadRequest request) {
        String sourceLocation = resolveSourceLocation(request);
        String fileExtension = FilenameUtils.getExtension(sourceLocation);
        if (fileExtension == null) {
            return false;
        }
        fileExtension = fileExtension.toLowerCase(Locale.ROOT);
        return ADVANCED_SUPPORTED_EXTENSIONS.contains(fileExtension);
    }

    private String resolveSourceLocation(ExcelReadRequest request) {
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
