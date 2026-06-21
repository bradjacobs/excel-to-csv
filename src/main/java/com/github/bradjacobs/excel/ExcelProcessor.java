/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.api.ExcelWorkbookReader;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.engine.eventmodel.EventModelExcelReader;
import com.github.bradjacobs.excel.engine.usermodel.StandardExcelReader;
import com.github.bradjacobs.excel.model.SheetContent;
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
    private static final List<String> EVENT_SUPPORTED_EXTENSIONS
            = List.of("xlsx", "xlsb", "xlsm", "xltx", "xltm");

    private final boolean useEventReader;
    private final ExcelWorkbookReader standardExcelWorkbookReader;
    private final ExcelWorkbookReader eventExcelWorkbookReader;

    private ExcelProcessor(Builder builder) {
        this.useEventReader = builder.useEventReader;
        this.standardExcelWorkbookReader = builder.standardBuilder.build();
        this.eventExcelWorkbookReader = builder.eventBuilder.build();
    }

    @Override
    public List<SheetContent> readSheets(ExcelReadRequest request) throws IOException {
        Validate.isTrue(request != null, "Request cannot be null");
        return getSheetReaderForRequest(request).readSheets(request);
    }

    private ExcelWorkbookReader getSheetReaderForRequest(ExcelReadRequest request) {
        return shouldUseEventReader(request)
                ? this.eventExcelWorkbookReader
                : this.standardExcelWorkbookReader;
    }

    private boolean shouldUseEventReader(ExcelReadRequest request) {
        return this.useEventReader && hasEventSupportedExtension(request);
    }

    private boolean hasEventSupportedExtension(ExcelReadRequest request) {
        String sourceLocation = resolveSourceLocation(request);
        String fileExtension = FilenameUtils.getExtension(sourceLocation);
        if (fileExtension == null) {
            return false;
        }
        fileExtension = fileExtension.toLowerCase(Locale.ROOT);
        return EVENT_SUPPORTED_EXTENSIONS.contains(fileExtension);
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

    // Special Builder that acts as composite builder for both standard and event readers.
    public static class Builder implements SheetConfig.ConfigBuilder<ExcelProcessor, Builder> {
        private boolean useEventReader = true;
        private final StandardExcelReader.Builder standardBuilder = StandardExcelReader.builder();
        private final EventModelExcelReader.Builder eventBuilder = EventModelExcelReader.builder();

        /**
         * Toggle preference for using the event reader.
         * Note this value will be ignored when trying to read old '.xls' files.
         * (typically only used for testing convenience)
         * @param useEventReader true to use the event reader... "if possible"
         * @return this builder
         */
        public Builder useEventReader(boolean useEventReader) {
            this.useEventReader = useEventReader;
            return this;
        }

        @Override
        public Builder disableAllSanitation() {
            standardBuilder.disableAllSanitation();
            eventBuilder.disableAllSanitation();
            return this;
        }

        @Override
        public Builder trimStringValues(boolean trimStringValues) {
            standardBuilder.trimStringValues(trimStringValues);
            eventBuilder.trimStringValues(trimStringValues);
            return this;
        }

        @Override
        public Builder skipBlankRows(boolean skipBlankRows) {
            standardBuilder.skipBlankRows(skipBlankRows);
            eventBuilder.skipBlankRows(skipBlankRows);
            return this;
        }

        @Override
        public Builder skipBlankColumns(boolean skipBlankColumns) {
            standardBuilder.skipBlankColumns(skipBlankColumns);
            eventBuilder.skipBlankColumns(skipBlankColumns);
            return this;
        }

        @Override
        public Builder skipHiddenCells(boolean skipHiddenCells) {
            standardBuilder.skipHiddenCells(skipHiddenCells);
            eventBuilder.skipHiddenCells(skipHiddenCells);
            return this;
        }

        @Override
        public Builder sanitizeSpaces(boolean sanitizeSpaces) {
            standardBuilder.sanitizeSpaces(sanitizeSpaces);
            eventBuilder.sanitizeSpaces(sanitizeSpaces);
            return this;
        }

        @Override
        public Builder sanitizeQuotes(boolean sanitizeQuotes) {
            standardBuilder.sanitizeQuotes(sanitizeQuotes);
            eventBuilder.sanitizeQuotes(sanitizeQuotes);
            return this;
        }

        @Override
        public Builder sanitizeDiacritics(boolean sanitizeDiacritics) {
            standardBuilder.sanitizeDiacritics(sanitizeDiacritics);
            eventBuilder.sanitizeDiacritics(sanitizeDiacritics);
            return this;
        }

        @Override
        public Builder sanitizeDashes(boolean sanitizeDashes) {
            standardBuilder.sanitizeDashes(sanitizeDashes);
            eventBuilder.sanitizeDashes(sanitizeDashes);
            return this;
        }

        @Override
        public ExcelProcessor build() {
            return new ExcelProcessor(this);
        }
    }
}
