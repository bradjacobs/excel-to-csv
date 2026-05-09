/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import com.github.bradjacobs.excel.io.InputStreamGenerator;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Path;
import java.util.Collection;

public class ExcelSheetReadRequest {
    private static final String DEFAULT_PASSWORD = null;
    private static final int DEFAULT_SHEET_INDEX = 0;
    // by default only process the first sheet in Excel file.
    private static final SheetSelector DEFAULT_SHEET_SELECTOR =
            new ByIndexSheetSelector(DEFAULT_SHEET_INDEX);

    private final Path path;
    private final URL url;
    private final SheetSelector sheetSelector;
    private final String password;

    private ExcelSheetReadRequest(Builder builder) {
        this.path = builder.path;
        this.url = builder.url;
        this.sheetSelector = builder.sheetSelector;
        this.password = builder.password;
    }

    // Factory methods for readability
    public static Builder from(Path path) {
        return new Builder(path, null);
    }

    public static Builder from(File file) {
        return from(file != null ? file.toPath() : null);
    }

    public static Builder from(URL url) {
        return new Builder(null, url);
    }

    // ===== Builder =====
    public static class Builder {
        private final Path path;
        private final URL url;
        private SheetSelector sheetSelector = DEFAULT_SHEET_SELECTOR;
        private String password = DEFAULT_PASSWORD;

        // Constructor enforces at least one required field
        private Builder(Path path, URL url) {
            Validate.isTrue(path != null || url != null,
                    "Either file path or url must be provided");
            this.path = path;
            this.url = url;
        }

        // Optional setters
        public Builder bySheetName(String name) {
            return bySheetNames(name);
        }
        public Builder bySheetNames(String ... names) {
            return sheetSelector(new ByNameSheetSelector(names));
        }
        public Builder bySheetNames(Collection<String> names) {
            return sheetSelector(new ByNameSheetSelector(names));
        }

        public Builder bySheetIndex(int index) {
            return bySheetIndexes(index);
        }
        public Builder bySheetIndexes(int ... indexes) {
            return sheetSelector(new ByIndexSheetSelector(indexes));
        }
        public Builder bySheetIndexes(Collection<Integer> indexes) {
            return sheetSelector(new ByIndexSheetSelector(indexes));
        }

        public Builder allSheets() {
            return sheetSelector(new AllSheetsSelector());
        }

        public Builder sheetSelector(SheetSelector sheetSelector) {
            this.sheetSelector = sheetSelector != null ? sheetSelector : DEFAULT_SHEET_SELECTOR;
            return this;
        }

        public Builder password(String password) {
            this.password = StringUtils.isNotEmpty(password) ? password : null;
            return this;
        }

        // Build method
        public ExcelSheetReadRequest build() {
            return new ExcelSheetReadRequest(this);
        }
    }

    // ===== Getters =====
    public Path getPath() {
        return path;
    }

    public URL getUrl() {
        return url;
    }

    public SheetSelector getSheetSelector() {
        return sheetSelector;
    }

    public String getPassword() {
        return password;
    }

    public InputStream getSourceInputStream() throws IOException {
        InputStreamGenerator inputStreamGenerator = new InputStreamGenerator();
        if (path != null) {
            return inputStreamGenerator.getInputStream(path);
        }
        else {
            return inputStreamGenerator.getInputStream(url);
        }
    }
}
