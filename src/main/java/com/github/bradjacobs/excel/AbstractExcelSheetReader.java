/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.io.InputStreamGenerator;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.SanitizeType.DASHES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.SanitizeType.QUOTES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.SanitizeType.SPACES;

abstract public class AbstractExcelSheetReader implements ExcelSheetReader {

    protected static final InputStreamGenerator inputStreamGenerator = new InputStreamGenerator();
    private static final String DEFAULT_PASSWORD = null;

    public AbstractExcelSheetReader() {
        // override the internal POI utils size limit to allow for 'bigger Excel files'
        //   (as of POI version 5.2.0 the default value is 100_000_000)
        org.apache.poi.util.IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
    }

    // Variations of reading Excel Sheet via sheet Index.
    @Override
    public String[][] readExcelSheetData(File excelFile, int sheetIndex) throws IOException {
        return readExcelSheetData(excelFile, sheetIndex, DEFAULT_PASSWORD);
    }

    @Override
    public String[][] readExcelSheetData(Path excelFile, int sheetIndex) throws IOException {
        return readExcelSheetData(excelFile, sheetIndex, DEFAULT_PASSWORD);
    }

    @Override
    public String[][] readExcelSheetData(URL excelFileUrl, int sheetIndex) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFileUrl), sheetIndex, DEFAULT_PASSWORD);
    }

    @Override
    public String[][] readExcelSheetData(InputStream inputStream, int sheetIndex) throws IOException {
        return readExcelSheetData(inputStream, sheetIndex, DEFAULT_PASSWORD);
    }

    @Override
    public String[][] readExcelSheetData(File excelFile, int sheetIndex, String password) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFile), sheetIndex, password);
    }

    @Override
    public String[][] readExcelSheetData(Path excelFile, int sheetIndex, String password) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFile), sheetIndex, password);
    }

    // Subclasses implement the core extraction logic:
    @Override
    public abstract String[][] readExcelSheetData(InputStream inputStream, int sheetIndex, String password) throws IOException;


    // Variations of reading Excel Sheet via sheet name.
    @Override
    public String[][] readExcelSheetData(File excelFile, String sheetName) throws IOException {
        return readExcelSheetData(excelFile, sheetName, DEFAULT_PASSWORD);
    }

    @Override
    public String[][] readExcelSheetData(Path excelFile, String sheetName) throws IOException {
        return readExcelSheetData(excelFile, sheetName, DEFAULT_PASSWORD);
    }

    @Override
    public String[][] readExcelSheetData(URL excelFileUrl, String sheetName) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFileUrl), sheetName, null);
    }

    @Override
    public String[][] readExcelSheetData(InputStream inputStream, String sheetName) throws IOException {
        return readExcelSheetData(inputStream, sheetName, DEFAULT_PASSWORD);
    }

    @Override
    public String[][] readExcelSheetData(File excelFile, String sheetName, String password) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFile), sheetName, password);
    }

    @Override
    public String[][] readExcelSheetData(Path excelFile, String sheetName, String password) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFile), sheetName, password);
    }

    // Subclasses implement the core extraction logic:
    @Override
    public abstract String[][] readExcelSheetData(InputStream inputStream, String sheetName, String password) throws IOException;


    // below is common code for sheet configuration builder
    abstract public static class AbstractSheetConfigBuilder<T, B extends AbstractSheetConfigBuilder<T, B>> {
        protected boolean autoTrim = true;
        protected boolean removeBlankRows = false;
        protected boolean removeBlankColumns = false;
        protected boolean removeInvisibleCells = false;
        protected Set<SpecialCharacterSanitizer.SanitizeType> sanitizeTypes
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
         * Whether to remove any blank rows.
         * @param removeBlankRows (defaults to false)
         */
        public B removeBlankRows(boolean removeBlankRows) {
            this.removeBlankRows = removeBlankRows;
            return self();
        }

        /**
         * Whether to remove any blank columns.
         * @param removeBlankColumns (defaults to false)
         */
        public B removeBlankColumns(boolean removeBlankColumns) {
            this.removeBlankColumns = removeBlankColumns;
            return self();
        }

        /**
         * Whether to prune out invisible cells.
         *   invisible = cellHeight = 0 or cellWidth = 0
         * @param removeInvisibleCells (defaults to false)
         */
        public B removeInvisibleCells(boolean removeInvisibleCells) {
            this.removeInvisibleCells = removeInvisibleCells;
            return self();
        }

        public B sanitizeSpaces(boolean sanitizeSpaces) {
            return setSanitizeFlag(SPACES, sanitizeSpaces);
        }

        public B sanitizeQuotes(boolean sanitizeQuotes) {
            return setSanitizeFlag(QUOTES, sanitizeQuotes);
        }

        public B sanitizeDiacritics(boolean sanitizeDiacritics) {
            return setSanitizeFlag(BASIC_DIACRITICS, sanitizeDiacritics);
        }

        public B sanitizeDashes(boolean sanitizeDashes) {
            return setSanitizeFlag(DASHES, sanitizeDashes);
        }

        private B setSanitizeFlag(SpecialCharacterSanitizer.SanitizeType flag, boolean shouldAdd) {
            if (shouldAdd) {
                this.sanitizeTypes.add(flag);
            }
            else {
                this.sanitizeTypes.remove(flag);
            }
            return self();
        }

        protected SheetConfig buildConfig() {
            return new SheetConfig(
                    removeBlankRows,
                    removeBlankColumns,
                    removeInvisibleCells,
                    autoTrim,
                    sanitizeTypes);
        }

        abstract public T build();
    }
}
