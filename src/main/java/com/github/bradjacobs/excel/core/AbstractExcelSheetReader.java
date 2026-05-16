/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import com.github.bradjacobs.excel.api.ExcelSheetReader;
import com.github.bradjacobs.excel.config.SanitizeType;
import com.github.bradjacobs.excel.config.SheetConfig;
import org.apache.commons.lang3.Validate;

import java.util.HashSet;
import java.util.Set;

import static com.github.bradjacobs.excel.config.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.config.SanitizeType.DASHES;
import static com.github.bradjacobs.excel.config.SanitizeType.QUOTES;
import static com.github.bradjacobs.excel.config.SanitizeType.SPACES;


// TODO - this abstract class might be removed.
//   its usefulness has now become very limited.
abstract public class AbstractExcelSheetReader implements ExcelSheetReader {

    protected final SheetConfig sheetConfig;

    public AbstractExcelSheetReader(SheetConfig sheetConfig) {
        Validate.isTrue(sheetConfig != null, "SheetConfig cannot be null.");
        this.sheetConfig = sheetConfig;

        // override the internal POI utils size limit to allow for 'bigger Excel files'
        //   (as of POI version 5.2.0 the default value is 100_000_000)
        org.apache.poi.util.IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
    }

    // below is common code for sheet configuration builder
    abstract public static class AbstractSheetConfigBuilder<T, B extends AbstractSheetConfigBuilder<T, B>> {
        protected boolean trimStringValues = true;
        protected boolean skipBlankRows = false;
        protected boolean skipBlankColumns = false;
        protected boolean skipHiddenCells = false;
        protected final Set<SanitizeType> sanitizeTypes
                = new HashSet<>(SpecialCharacterSanitizer.DEFAULT_FLAGS);

        protected abstract B self();

        /**
         * Disable all default sanitation settings.
         * Intended for when you 'know' the input does not
         * require any 'trimming' or handling of special Unicode characters.
         * Can result in a 'mild' performance improvement. (est. ~ 10%)
         */
        public B disableAllSanitation() {
            this.sanitizeTypes.clear();
            this.trimStringValues = false;
            return self();
        }

        /**
         * Whether to trim whitespace on cell values
         * @param trimStringValues (defaults to true)
         */
        public B trimStringValues(boolean trimStringValues) {
            this.trimStringValues = trimStringValues;
            return self();
        }

        /**
         * Whether to skip any blank rows.
         * @param skipBlankRows (defaults to false)
         */
        public B skipBlankRows(boolean skipBlankRows) {
            this.skipBlankRows = skipBlankRows;
            return self();
        }

        /**
         * Whether to skip any blank columns.
         * @param skipBlankColumns (defaults to false)
         */
        public B skipBlankColumns(boolean skipBlankColumns) {
            this.skipBlankColumns = skipBlankColumns;
            return self();
        }

        /**
         * Whether to prune out hidden cells.
         *   hidden = cellHeight = 0 or cellWidth = 0
         * @param skipHiddenCells (defaults to false)
         */
        public B skipHiddenCells(boolean skipHiddenCells) {
            this.skipHiddenCells = skipHiddenCells;
            return self();
        }

        public B sanitizeSpaces(boolean sanitizeSpaces) {
            return setSanitizeType(SPACES, sanitizeSpaces);
        }

        public B sanitizeQuotes(boolean sanitizeQuotes) {
            return setSanitizeType(QUOTES, sanitizeQuotes);
        }

        public B sanitizeDiacritics(boolean sanitizeDiacritics) {
            return setSanitizeType(BASIC_DIACRITICS, sanitizeDiacritics);
        }

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

        public SheetConfig buildConfig() {
            return new SheetConfig(
                    skipBlankRows,
                    skipBlankColumns,
                    skipHiddenCells,
                    trimStringValues,
                    sanitizeTypes);
        }

        abstract public T build();
    }
}
