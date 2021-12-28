/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package bwj.util.excel;

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
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

/**
 * Simple class that reads an Excel Worksheet
 *   and will produce a CSV-equivalent
 */
// todo: javadocs
public class ExcelReader
{
    private static final Set<String> VALID_URL_SCHEMES =
            new HashSet<>(Arrays.asList("http", "https", "ftp", "file"));
    private static final int CONNECTION_TIMEOUT = 20000;

    // some websites require a userAgent value set.
    //    side:  seen a case where a userAgent with substring 'java' would fail  (empirical evidence)
    private static final String USER_AGENT_VALUE = "jclient/" + System.getProperty("java.version");


    private final int sheetIndex;
    private final String sheetName;
    private final MatrixToCsvTextConverter matrixToCsvTextConverter;
    private final ExcelSheetReader excelSheetToCsvConverter;

    private final InputStreamGenerator inputStreamGenerator;

    private ExcelReader(Builder builder)
    {
        this.sheetIndex = builder.sheetIndex;
        this.sheetName = builder.sheetName;
        this.matrixToCsvTextConverter = new MatrixToCsvTextConverter(builder.quoteMode);
        this.excelSheetToCsvConverter = new ExcelSheetReader(builder.skipEmptyRows);

        this.inputStreamGenerator = new InputStreamGenerator();
        this.inputStreamGenerator.setUseGzip(builder.gzipEnabled);
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
        return convertToDataMatrix(  inputStreamGenerator.getInputStream(excelUrl) );
    }

    private String[][] convertToDataMatrix(InputStream inputStream) throws IOException {
        try (Workbook wb = WorkbookFactory.create(inputStream)) {
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
        private boolean gzipEnabled = true;

        private Builder() {}

        /**
         * Set with sheet of Excel file to read (defaults to '0', i.e. the first sheet)
         * @param sheetIndex (0-based index of which sheet in excel file to convert)
         */
        public Builder setSheetIndex(int sheetIndex) {
            this.sheetIndex = sheetIndex;
            return this;
        }

        /**
         * Optionally can provide a sheet name instead of an index
         *  (if sheetName is set then sheetIndex is ignored)
         * @param sheetName name of Excel sheet (case-sensitive)
         */
        public Builder setSheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        /**
         * Whether to skip any empty rows.
         * @param skipEmptyRows (defaults to false)
         */
        public Builder setSkipEmptyRows(boolean skipEmptyRows) {
            this.skipEmptyRows = skipEmptyRows;
            return this;
        }

        public Builder setGzipEnabled(boolean gzipEnabled) {
            this.gzipEnabled = gzipEnabled;
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
            this.quoteMode = quoteMode;
            return this;
        }

        public ExcelReader build() {
            validateInputs();
            return new ExcelReader(this);
        }

        private void validateInputs() {
            if (this.sheetIndex < 0) {
                throw new IllegalArgumentException("SheetIndex cannot be negative");
            }
            if (this.quoteMode == null) {
                throw new IllegalArgumentException("Cannot set quoteMode to null");
            }
        }
    }
}
