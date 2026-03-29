/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.config;

import java.util.Collections;
import java.util.Set;

public class SheetConfig {

    private final boolean skipBlankRows;
    private final boolean skipBlankColumns;
    private final boolean skipInvisibleCells;
    private final boolean autoTrim;
    private final Set<SanitizeType> sanitizeTypes;

    public SheetConfig(
            boolean skipBlankRows,
            boolean skipBlankColumns,
            boolean skipInvisibleCells,
            boolean autoTrim,
            Set<SanitizeType> sanitizeTypes) {
        this.skipBlankRows = skipBlankRows;
        this.skipBlankColumns = skipBlankColumns;
        this.skipInvisibleCells = skipInvisibleCells;
        this.autoTrim = autoTrim;
        this.sanitizeTypes = Collections.unmodifiableSet(sanitizeTypes);
    }

    public boolean skipBlankRows() {
        return skipBlankRows;
    }

    public boolean skipBlankColumns() {
        return skipBlankColumns;
    }

    public boolean skipInvisibleCells() {
        return skipInvisibleCells;
    }

    public boolean isAutoTrim() {
        return autoTrim;
    }

    public Set<SanitizeType> getCharSanitizeFlags() {
        return sanitizeTypes;
    }
}
