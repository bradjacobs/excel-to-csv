/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.charset.StandardCharsets;

/**
 * Simple class that reads an Excel Worksheet
 *   and will produce a CSV-equivalent
 */
// todo: javadocs
public class ExcelReader {
    private final int sheetIndex;
    private final String sheetName;
    private final String password; // 'null' == no password
    private final MatrixToCsvTextConverter matrixToCsvTextConverter;
    private final ExcelSheetReader excelSheetToCsvConverter;

    private final InputStreamGenerator inputStreamGenerator;

    private ExcelReader(Builder builder) {
        this.sheetIndex = builder.sheetIndex;
        this.sheetName = builder.sheetName;
        this.password = builder.password;
        this.matrixToCsvTextConverter = new MatrixToCsvTextConverter(builder.quoteMode);
        this.excelSheetToCsvConverter = new ExcelSheetReader(builder.skipEmptyRows);
        this.inputStreamGenerator = new InputStreamGenerator();
    }

    public void convertToCsvFile(File excelFile, File outputFile) throws IOException {
        writeCsvToFile( convertToCsvText(excelFile), outputFile);
    }
    public void convertToCsvFile(URL excelUrl, File outputFile) throws IOException {
        writeCsvToFile( convertToCsvText(excelUrl), outputFile);
    }

    public String convertToCsvText(File excelFile) throws IOException {
        return matrixToCsvTextConverter.createCsvText( convertToDataMatrix(excelFile) );
    }
    public String convertToCsvText(URL excelUrl) throws IOException {
        return matrixToCsvTextConverter.createCsvText( convertToDataMatrix(excelUrl) );
    }

    public String[][] convertToDataMatrix(File excelFile) throws IOException {
        return convertToDataMatrix( inputStreamGenerator.getInputStream(excelFile) );
    }
    public String[][] convertToDataMatrix(URL excelUrl) throws IOException {
        return convertToDataMatrix( inputStreamGenerator.getInputStream(excelUrl) );
    }

    private String[][] convertToDataMatrix(InputStream inputStream) throws IOException {
        try (Workbook wb = WorkbookFactory.create(inputStream, password)) {
            Sheet sheet = getSheet(wb);
            return excelSheetToCsvConverter.convertToCsvData(sheet);
        }
        finally {
            // ensure inputStream closed b/c Workbook/WorkbookFactory might not do it.
            inputStream.close();
        }
    }

    private Sheet getSheet(Workbook wb) {
        Sheet returnSheet;
        if (StringUtils.isNotEmpty(this.sheetName)) {
            returnSheet = wb.getSheet(this.sheetName);
            if (returnSheet == null) {
                throw new IllegalArgumentException(String.format("Unable to find sheet with name: %s", this.sheetName));
            }
        }
        else {
            returnSheet = wb.getSheetAt(this.sheetIndex);
        }
        return returnSheet;
    }

    private void writeCsvToFile(String csvString, File outputFile) throws IOException {
        if (outputFile == null) {
            throw new IllegalArgumentException("Must supply outputFile location to save CSV data.");
        }
        else if (outputFile.isDirectory()) {
            throw new IllegalArgumentException("The outputFile cannot be an existing directory.");
        }
        FileUtils.writeStringToFile(outputFile, csvString, StandardCharsets.UTF_8);
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder {
        private int sheetIndex = 0;    // default to the first tab
        private String sheetName = ""; // optionally provide a specific sheet name
        private boolean skipEmptyRows = true; // default will skip any empty lines
        private QuoteMode quoteMode = QuoteMode.NORMAL;
        private String password = null;

        private Builder() {}

        /**
         * Set with sheet of Excel file to read (defaults to '0', i.e. the first sheet)
         * @param sheetIndex (0-based index of which sheet in Excel file to convert)
         */
        public Builder setSheetIndex(int sheetIndex) {
            if (sheetIndex < 0) {
                throw new IllegalArgumentException("SheetIndex cannot be negative");
            }
            this.sheetIndex = sheetIndex;
            return this;
        }

        /**
         * Optionally can provide a sheet name instead of an index
         *  (if sheetName is set then sheetIndex is ignored)
         * @param sheetName name of Excel sheet
         */
        public Builder setSheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        /**
         * Whether to skip any empty rows.
         * @param skipEmptyRows (defaults to true)
         */
        public Builder setSkipEmptyRows(boolean skipEmptyRows) {
            this.skipEmptyRows = skipEmptyRows;
            return this;
        }

        /**
         * Set how to handle quote/escaping string values to be CSV-compliant
         * @param quoteMode
         *  ALWAYS:  surround all values with quotes
         *  NORMAL:  add quotes around most values that contain non-alphanumeric (roughly similar to Jackson CsvMapper)
         *  LENIENT: add quotes around values that only really 'need' it to adhere to valid CSV (roughly similar to Excel 'save-as' CSV)
         *  NEVER:   never add quotes to any values.
         */
        public Builder setQuoteMode(QuoteMode quoteMode) {
            if (quoteMode == null) {
                throw new IllegalArgumentException("Cannot set quoteMode to null");
            }
            this.quoteMode = quoteMode;
            return this;
        }

        public Builder setPassword(String password) {
            this.password = password != null && password.isEmpty() ? null : password;
            return this;
        }

        public ExcelReader build() {
            return new ExcelReader(this);
        }
    }
}
