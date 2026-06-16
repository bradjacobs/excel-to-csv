/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.config;

import com.github.bradjacobs.excel.sanitize.SanitizeType;
import com.github.bradjacobs.excel.sanitize.SpecialCharacterSanitizer;

import java.util.HashSet;
import java.util.Set;

import static com.github.bradjacobs.excel.sanitize.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.sanitize.SanitizeType.DASHES;
import static com.github.bradjacobs.excel.sanitize.SanitizeType.QUOTES;
import static com.github.bradjacobs.excel.sanitize.SanitizeType.SPACES;

public class SheetConfig {

    private final boolean skipBlankRows;
    private final boolean skipBlankColumns;
    private final boolean skipHiddenCells;
    private final boolean trimStringValues;
    private final Set<SanitizeType> sanitizeTypes;

    private SheetConfig(
            boolean skipBlankRows,
            boolean skipBlankColumns,
            boolean skipHiddenCells,
            boolean trimStringValues,
            Set<SanitizeType> sanitizeTypes) {
        this.skipBlankRows = skipBlankRows;
        this.skipBlankColumns = skipBlankColumns;
        this.skipHiddenCells = skipHiddenCells;
        this.trimStringValues = trimStringValues;
        this.sanitizeTypes = Set.copyOf(sanitizeTypes);
    }

    public boolean skipBlankRows() {
        return skipBlankRows;
    }

    public boolean skipBlankColumns() {
        return skipBlankColumns;
    }

    public boolean skipHiddenCells() {
        return skipHiddenCells;
    }

    public boolean trimStringValues() {
        return trimStringValues;
    }

    public Set<SanitizeType> getCharSanitizeFlags() {
        return sanitizeTypes;
    }

    // Builder...


    public interface ConfigBuilder<T, B extends ConfigBuilder<T, B>> {
        /**
         * Disable all default sanitation settings.
         * Intended for when you 'know' the input does not
         * require any 'trimming' or handling of special Unicode characters.
         * Can result in a 'mild' performance improvement.
         */
        B disableAllSanitation();

        /**
         * Whether to trim whitespace on cell values
         * @param trimStringValues (defaults to true)
         */
        B trimStringValues(boolean trimStringValues);

        /**
         * Whether to skip any blank rows.
         * @param skipBlankRows (defaults to false)
         */
        B skipBlankRows(boolean skipBlankRows);

        /**
         * Whether to skip any blank columns.
         * @param skipBlankColumns (defaults to false)
         */
        B skipBlankColumns(boolean skipBlankColumns);

        /**
         * Whether to prune out hidden cells.
         * Hidden = cellHeight = 0 or cellWidth = 0
         * @param skipHiddenCells (defaults to false)
         */
        B skipHiddenCells(boolean skipHiddenCells);

        B sanitizeSpaces(boolean sanitizeSpaces);
        B sanitizeQuotes(boolean sanitizeQuotes);
        B sanitizeDiacritics(boolean sanitizeDiacritics);
        B sanitizeDashes(boolean sanitizeDashes);

        T build();
    }

    // abstract builder class that can be used by the _ExcelReader classes as well.
    abstract public static class AbstractSheetConfigBuilder<T, B extends AbstractSheetConfigBuilder<T, B>> implements ConfigBuilder<T, B>{
        protected boolean trimStringValues = true;
        protected boolean skipBlankRows = false;
        protected boolean skipBlankColumns = false;
        protected boolean skipHiddenCells = false;
        protected final Set<SanitizeType> sanitizeTypes
                = new HashSet<>(SpecialCharacterSanitizer.DEFAULT_FLAGS);

        protected abstract B self();

        @Override
        public B disableAllSanitation() {
            this.sanitizeTypes.clear();
            this.trimStringValues = false;
            return self();
        }

        @Override
        public B trimStringValues(boolean trimStringValues) {
            this.trimStringValues = trimStringValues;
            return self();
        }

        @Override
        public B skipBlankRows(boolean skipBlankRows) {
            this.skipBlankRows = skipBlankRows;
            return self();
        }

        @Override
        public B skipBlankColumns(boolean skipBlankColumns) {
            this.skipBlankColumns = skipBlankColumns;
            return self();
        }

        @Override
        public B skipHiddenCells(boolean skipHiddenCells) {
            this.skipHiddenCells = skipHiddenCells;
            return self();
        }

        @Override
        public B sanitizeSpaces(boolean sanitizeSpaces) {
            return setSanitizeType(SPACES, sanitizeSpaces);
        }

        @Override
        public B sanitizeQuotes(boolean sanitizeQuotes) {
            return setSanitizeType(QUOTES, sanitizeQuotes);
        }

        @Override
        public B sanitizeDiacritics(boolean sanitizeDiacritics) {
            return setSanitizeType(BASIC_DIACRITICS, sanitizeDiacritics);
        }

        @Override
        public B sanitizeDashes(boolean sanitizeDashes) {
            return setSanitizeType(DASHES, sanitizeDashes);
        }

        private B setSanitizeType(SanitizeType type, boolean shouldAdd) {
            if (shouldAdd) {
                this.sanitizeTypes.add(type);
            }
            else {
                this.sanitizeTypes.remove(type);
            }
            return self();
        }

        protected SheetConfig buildConfig() {
            return new SheetConfig(
                    skipBlankRows,
                    skipBlankColumns,
                    skipHiddenCells,
                    trimStringValues,
                    sanitizeTypes);
        }

        abstract public T build();
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder extends AbstractSheetConfigBuilder<SheetConfig, Builder> {
        @Override
        protected Builder self() {
            return this;
        }

        @Override
        public SheetConfig build() {
            return this.buildConfig();
        }
    }
}
