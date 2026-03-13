package com.github.bradjacobs.excel.config;

import com.github.bradjacobs.excel.SpecialCharacterSanitizer;

import java.util.Collections;
import java.util.Set;

public class SheetConfig {

    private final boolean removeBlankRows;
    private final boolean removeBlankColumns;
    private final boolean removeInvisibleCells;
    private final boolean autoTrim;
    private final Set<SpecialCharacterSanitizer.SanitizeType> sanitizeTypes;

    public SheetConfig(
            boolean removeBlankRows,
            boolean removeBlankColumns,
            boolean removeInvisibleCells,
            boolean autoTrim,
            Set<SpecialCharacterSanitizer.SanitizeType> sanitizeTypes) {
        this.removeBlankRows = removeBlankRows;
        this.removeBlankColumns = removeBlankColumns;
        this.removeInvisibleCells = removeInvisibleCells;
        this.autoTrim = autoTrim;
        this.sanitizeTypes = Collections.unmodifiableSet(sanitizeTypes);
    }

    public boolean isRemoveBlankRows() {
        return removeBlankRows;
    }

    public boolean isRemoveBlankColumns() {
        return removeBlankColumns;
    }

    public boolean isRemoveInvisibleCells() {
        return removeInvisibleCells;
    }

    public boolean isAutoTrim() {
        return autoTrim;
    }

    public Set<SpecialCharacterSanitizer.SanitizeType> getCharSanitizeFlags() {
        return sanitizeTypes;
    }
}
