/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.QUOTES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.SPACES;

/**
 * Simple class that reads an Excel Worksheet
 *   and will produce a CSV-equivalent
 */
// todo: javadocs
public class ExcelReader {
    private final int sheetIndex;
    private final String sheetName;
    private final String password; // 'null' == no password
    private final boolean saveUnicodeFileWithBom;
    private final MatrixToCsvTextConverter matrixToCsvTextConverter;
    private final ExcelSheetReader excelSheetReader;

    private static final Set<String> ALLOWED_OUTPUT_FILE_EXTENSIONS = new HashSet<>(Arrays.asList("csv", "txt", ""));
    private static final String BOM = "\uFEFF";

    private final InputStreamGenerator inputStreamGenerator;

    private ExcelReader(Builder builder) {
        this.sheetIndex = builder.sheetIndex;
        this.sheetName = builder.sheetName;
        this.password = builder.password;
        this.saveUnicodeFileWithBom = builder.saveUnicodeFileWithBom;
        this.matrixToCsvTextConverter = new MatrixToCsvTextConverter(builder.quoteMode);
        this.excelSheetReader = new ExcelSheetReader(
                builder.skipEmptyRows, builder.charSanitizeFlags);
        this.inputStreamGenerator = new InputStreamGenerator();

        // override the internal POI utils size limit to allow for 'bigger Excel files'
        //   (as of POI version 5.2.0 the default value is 100_000_000)
        org.apache.poi.util.IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
    }

    public void convertToCsvFile(Path excelFile, Path outputFile) throws IOException {
        validateOutputFileParameter(outputFile);
        writeCsvToFile( convertToCsvText(excelFile), outputFile);
    }
    public void convertToCsvFile(File excelFile, File outputFile) throws IOException {
        convertToCsvFile(fileToPath(excelFile), fileToPath(outputFile));
    }
    public void convertToCsvFile(URL excelUrl, Path outputFile) throws IOException {
        validateOutputFileParameter(outputFile);
        writeCsvToFile( convertToCsvText(excelUrl), outputFile);
    }
    public void convertToCsvFile(URL excelUrl, File outputFile) throws IOException {
        convertToCsvFile(excelUrl, fileToPath(outputFile));
    }

    public String convertToCsvText(Path excelFile) throws IOException {
        return matrixToCsvTextConverter.createCsvText( convertToDataMatrix(excelFile) );
    }
    public String convertToCsvText(File excelFile) throws IOException {
        return convertToCsvText(fileToPath(excelFile));
    }
    public String convertToCsvText(URL excelUrl) throws IOException {
        return matrixToCsvTextConverter.createCsvText( convertToDataMatrix(excelUrl) );
    }

    public String[][] convertToDataMatrix(Path excelFile) throws IOException {
        return convertToDataMatrix( inputStreamGenerator.getInputStream(excelFile) );
    }
    public String[][] convertToDataMatrix(File excelFile) throws IOException {
        return convertToDataMatrix(fileToPath(excelFile));
    }
    public String[][] convertToDataMatrix(URL excelUrl) throws IOException {
        return convertToDataMatrix( inputStreamGenerator.getInputStream(excelUrl) );
    }

    private String[][] convertToDataMatrix(InputStream inputStream) throws IOException {
        try (inputStream; Workbook wb = WorkbookFactory.create(inputStream, password)) {
            Sheet sheet = getSheet(wb);
            return excelSheetReader.convertToCsvData(sheet);
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

    /**
     * Write the CSV data string out to a file.
     * @param csvString CSV data
     * @param outputFile destination file.
     */
    private void writeCsvToFile(String csvString, Path outputFile) throws IOException {
        // prepend the 'bom' so that unicode characters will render correctly
        if (this.saveUnicodeFileWithBom && containsUnicode(csvString)) {
            csvString = BOM + csvString;
        }
        Files.writeString(outputFile, csvString, StandardCharsets.UTF_8);
    }

    /**
     * Some sanity checks on the outputFile parameter prior to doing the actual conversion
     *   (not exhaustive on all possible checks one could do)
     * @param outputFile the destination CSV output file.
     * @throws IllegalArgumentException if a problem was detected with the File object.
     */
    private void validateOutputFileParameter(Path outputFile) throws IllegalArgumentException {
        if (outputFile == null) {
            throw new IllegalArgumentException("Must supply outputFile location to save CSV data.");
        }
        else if (Files.isDirectory(outputFile)) {
            throw new IllegalArgumentException("The outputFile cannot be an existing directory.");
        }

        // convert to absolute path and continue checks...
        outputFile = outputFile.toAbsolutePath();

        // confirm output file has an allowed file extension
        String ext = FilenameUtils.getExtension(outputFile.toString());
        if (! ALLOWED_OUTPUT_FILE_EXTENSIONS.contains(ext.toLowerCase())) {
            throw new IllegalArgumentException(
                    String.format("Illegal outputFile extension '%s'.  Must be either 'csv', 'txt' or blank", ext));
        }

        Path parentDirectory = outputFile.getParent();
        if (parentDirectory == null || !Files.isDirectory(parentDirectory)) {
            throw new IllegalArgumentException("Attempted to save CSV output file in a non-existent directory: " + outputFile);
        }
    }

    private Path fileToPath(File file) {
        if (file == null) {
            return null;
        }
        try {
            return Paths.get(file.getAbsolutePath());
        }
        catch (InvalidPathException ex) {
            throw new IllegalArgumentException("The file path contains an illegal character: "
                    + file.getAbsolutePath());
        }
    }

    private boolean containsUnicode(String input) {
        int inputLength = input.length();
        for (int i = 0; i < inputLength; i++) {
            int codePoint = input.codePointAt(i);
            if (codePoint > 127) {
                return true;
            }
        }
        return false;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder {
        private static final CharSanitizeFlag[] DEFAULT_SANITIZE_FLAGS = { SPACES, QUOTES };

        private int sheetIndex = 0;    // default to the first tab
        private String sheetName = ""; // optionally provide a specific sheet name
        private boolean skipEmptyRows = false; // skip any empty lines when set
        private QuoteMode quoteMode = QuoteMode.NORMAL;
        private String password = null;
        private boolean saveUnicodeFileWithBom = true; // flag to write file with BOM if contains unicode.
        private final Set<CharSanitizeFlag> charSanitizeFlags = new HashSet<>(Arrays.asList(DEFAULT_SANITIZE_FLAGS));

        private Builder() {}

        /**
         * Set with sheet of Excel file to read (defaults to '0', i.e. the first sheet)
         * @param sheetIndex (0-based index of which sheet in Excel file to convert)
         */
        public Builder sheetIndex(int sheetIndex) {
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
        public Builder sheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        /**
         * Whether to skip any empty rows.
         * @param skipEmptyRows (defaults to true)
         */
        public Builder skipEmptyRows(boolean skipEmptyRows) {
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
        public Builder quoteMode(QuoteMode quoteMode) {
            if (quoteMode == null) {
                throw new IllegalArgumentException("Cannot set quoteMode to null");
            }
            this.quoteMode = quoteMode;
            return this;
        }

        /**
         * Define a password to open the Excel file (if needed)
         * @param password excel file password
         */
        public Builder password(String password) {
            // if user tries to set blank/empty string, then save as 'null'
            this.password = password != null && password.isEmpty() ? null : password;
            return this;
        }

        public Builder sanitizeWhitespace(boolean sanitizeWhitespace) {
            return setSanitizeFlag(SPACES, sanitizeWhitespace);
        }

        public Builder sanitizeQuotes(boolean sanitizeQuotes) {
            return setSanitizeFlag(QUOTES, sanitizeQuotes);
        }

        public Builder sanitizeDiacritics(boolean sanitizeDiacritics) {
            return setSanitizeFlag(BASIC_DIACRITICS, sanitizeDiacritics);
        }

        private Builder setSanitizeFlag(CharSanitizeFlag flag, boolean shouldAdd) {
            if (shouldAdd) { this.charSanitizeFlags.add(flag); }
            else { this.charSanitizeFlags.remove(flag); }
            return this;
        }

        /**
         * Use a BOM when writing output file if data contains 'unicode characters'
         * @param saveUnicodeFileWithBom (defaults to true)
         */
        public Builder saveUnicodeFileWithBom(boolean saveUnicodeFileWithBom) {
            this.saveUnicodeFileWithBom = saveUnicodeFileWithBom;
            return this;
        }

        public ExcelReader build() {
            return new ExcelReader(this);
        }
    }
}
