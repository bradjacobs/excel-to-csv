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
//   it's usefulness has now become very limited.
abstract public class AbstractExcelSheetReader implements ExcelSheetReader {

    private static final String SHEET_NOT_FOUND_MSG = "Excel sheet not found: '%s'";

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
        protected boolean autoTrim = true;
        protected boolean skipBlankRows = false;
        protected boolean skipBlankColumns = false;
        protected boolean skipInvisibleCells = false;
        protected final Set<SanitizeType> sanitizeTypes
                = new HashSet<>(SpecialCharacterSanitizer.DEFAULT_FLAGS);

        protected abstract B self();

        /**
         * Whether to trim whitespace on cell values
         * @param autoTrim (defaults to true)
         */
        public B autoTrim(boolean autoTrim) {
            this.autoTrim = autoTrim;
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
         * Whether to prune out invisible cells.
         *   invisible = cellHeight = 0 or cellWidth = 0
         * @param skipInvisibleCells (defaults to false)
         */
        public B skipInvisibleCells(boolean skipInvisibleCells) {
            this.skipInvisibleCells = skipInvisibleCells;
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
                    skipInvisibleCells,
                    autoTrim,
                    sanitizeTypes);
        }

        abstract public T build();
    }
}
