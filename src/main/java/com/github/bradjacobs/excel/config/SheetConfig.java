/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.config;

import java.util.Set;

public class SheetConfig {

    private final boolean skipBlankRows;
    private final boolean skipBlankColumns;
    private final boolean skipInvisibleCells;
    private final boolean trimStringValues;
    private final Set<SanitizeType> sanitizeTypes;

    public SheetConfig(
            boolean skipBlankRows,
            boolean skipBlankColumns,
            boolean skipInvisibleCells,
            boolean trimStringValues,
            Set<SanitizeType> sanitizeTypes) {
        this.skipBlankRows = skipBlankRows;
        this.skipBlankColumns = skipBlankColumns;
        this.skipInvisibleCells = skipInvisibleCells;
        this.trimStringValues = trimStringValues;
        this.sanitizeTypes = Set.copyOf(sanitizeTypes);
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

    public boolean trimStringValues() {
        return trimStringValues;
    }

    public Set<SanitizeType> getCharSanitizeFlags() {
        return sanitizeTypes;
    }
}
