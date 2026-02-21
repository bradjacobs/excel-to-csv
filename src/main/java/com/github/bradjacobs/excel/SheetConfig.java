package com.github.bradjacobs.excel;

import java.util.Collections;
import java.util.Set;

public class SheetConfig {

    private final boolean removeBlankRows;
    private final boolean removeBlankColumns;
    private final boolean removeInvisibleCells;
    private final boolean autoTrim;
    private final Set<SpecialCharacterSanitizer.CharSanitizeFlag> charSanitizeFlags;

    public SheetConfig(
            boolean removeBlankRows,
            boolean removeBlankColumns,
            boolean removeInvisibleCells,
            boolean autoTrim,
            Set<SpecialCharacterSanitizer.CharSanitizeFlag> charSanitizeFlags) {
        this.removeBlankRows = removeBlankRows;
        this.removeBlankColumns = removeBlankColumns;
        this.removeInvisibleCells = removeInvisibleCells;
        this.autoTrim = autoTrim;
        this.charSanitizeFlags = Collections.unmodifiableSet(charSanitizeFlags);
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

    public Set<SpecialCharacterSanitizer.CharSanitizeFlag> getCharSanitizeFlags() {
        return charSanitizeFlags;
    }
}
