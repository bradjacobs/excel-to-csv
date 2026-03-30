/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.api.ExcelSheetReader;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.AbstractExcelSheetReader.AbstractSheetConfigBuilder;
import com.github.bradjacobs.excel.csv.QuoteMode;
import com.github.bradjacobs.excel.io.InputStreamGenerator;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Path;

import static com.github.bradjacobs.excel.ExcelSheetReaderFactory.ReaderType.ADVANCED;
import static com.github.bradjacobs.excel.ExcelSheetReaderFactory.ReaderType.STANDARD;

/**
 * ExcelProcessor
 */
// TODO this is the 'next gen' of the ExcelReader
// TODO - in flux code
public class ExcelProcessor {
    private final int sheetIndex;
    private final String sheetName;
    private final String password; // 'null' == no password
    private final boolean useAdvancedReader;
    private final ExcelSheetReader excelSheetReader;
    private final ExcelSheetReader advancedExcelSheetReader;

    private final InputStreamGenerator inputStreamGenerator;

    private ExcelProcessor(Builder builder) {
        this.sheetIndex = builder.sheetIndex;
        this.sheetName = builder.sheetName;
        this.password = builder.password;
        this.useAdvancedReader = builder.useAdvancedReader;
        SheetConfig sheetConfig = builder.buildConfig();
        this.excelSheetReader = createSheetReader(STANDARD, sheetConfig);
        this.advancedExcelSheetReader = createSheetReader(ADVANCED, sheetConfig);
        this.inputStreamGenerator = new InputStreamGenerator();
    }

    public SheetData read(Path excelFile) throws IOException {
        return read( inputStreamGenerator.getInputStream(excelFile) );
    }
    public SheetData read(File excelFile) throws IOException {
        return read(fileToPath(excelFile));
    }
    public SheetData read(URL excelUrl) throws IOException {
        return read( inputStreamGenerator.getInputStream(excelUrl) );
    }

    private SheetData read(InputStream inputStream) throws IOException {
        // the inputStream should already be BufferedInputStream,
        //   but this is just an extra precaution.
        inputStream = FileMagic.prepareToCheckMagic(inputStream);

        final boolean advancedMode = this.useAdvancedReader && isOOXMLStream(inputStream);
        final boolean hasSheetName = StringUtils.isNotEmpty(this.sheetName);
        final ExcelSheetReader reader = advancedMode
                ? advancedExcelSheetReader
                : excelSheetReader;

        String[][] result = hasSheetName
                ? reader.readExcelSheetData(inputStream, this.sheetName, this.password)
                : reader.readExcelSheetData(inputStream, this.sheetIndex, this.password);

        return new SheetData("-sheetName-", result);
    }

    private static ExcelSheetReader createSheetReader(
            ExcelSheetReaderFactory.ReaderType readerType,
            SheetConfig sheetConfig) {
        return ExcelSheetReaderFactory.create(readerType, sheetConfig);
    }

    private Path fileToPath(File file) {
        return (file != null ? file.toPath() : null);
    }

    private boolean isOOXMLStream(InputStream inputStream) throws IOException {
        FileMagic fileMagic = FileMagic.valueOf(inputStream);
        return FileMagic.OOXML == fileMagic;

        // TODO
        //   add check if file is OLE2 + password protected.
        //   (original code to check had bug and was removed)
    }

    public static Builder builder() {
        return new Builder();
    }

    // this builder extends abstract class to allow any of the
    //   AbstractSheetConfigBuilder values to be set on this Builder as well.
    public static class Builder extends AbstractSheetConfigBuilder<ExcelProcessor, Builder> {
        private int sheetIndex = 0; // default to the first tab
        private String sheetName = ""; // optionally provide a specific sheet name
        private QuoteMode quoteMode = QuoteMode.NORMAL;
        private String password = null;
        private boolean useAdvancedReader = true; // use the advanced reader (if possible)

        @Override
        protected Builder self() {
            return this;
        }

        /**
         * Set with sheet of Excel file to read (defaults to '0', i.e. the first sheet)
         * @param sheetIndex (0-based index of which sheet in Excel file to convert)
         */
        public Builder sheetIndex(int sheetIndex) {
            Validate.isTrue(sheetIndex >= 0, "SheetIndex cannot be negative");
            this.sheetIndex = sheetIndex;
            return this;
        }

        /**
         * Optionally can provide a sheet name instead of an index
         *  (if sheetName is set then sheetIndex is ignored)
         * @param sheetName name of Excel sheet
         */
        public Builder sheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        /**
         * Set how to handle quote/escaping string values to be CSV-compliant
         * @param quoteMode
         *  ALWAYS:  surround all values with quotes
         *  NORMAL:  add quotes around most values that contain non-alphanumeric (roughly similar to Jackson CsvMapper)
         *  MINIMAL: add quotes around values that only really 'need' it to adhere to valid CSV (roughly similar to Excel 'save-as' CSV)
         *  NEVER:   never add quotes to any values.
         */
        public Builder quoteMode(QuoteMode quoteMode) {
            Validate.isTrue(quoteMode != null, "Cannot set quoteMode to null");
            this.quoteMode = quoteMode;
            return this;
        }

        /**
         * Define a password to open the Excel file (if needed)
         * @param password excel fExcelReader.Buildeile password
         */
        public Builder password(String password) {
            // if user tries to set blank/empty string, then save as 'null'
            this.password = StringUtils.isNotEmpty(password) ? password : null;
            return this;
        }

        /**
         * Toggle using the advanced reader
         *   (typically only used for testing convenience)
         */
        public Builder useAdvancedReader(boolean useAdvancedReader) {
            this.useAdvancedReader = useAdvancedReader;
            return this;
        }

        @Override
        public ExcelProcessor build() {
            return new ExcelProcessor(this);
        }
    }
}
