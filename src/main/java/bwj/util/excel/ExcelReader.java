/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package bwj.util.excel;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.nio.charset.StandardCharsets;

/**
 * Simple class that reads a Worksheet from an Excel file and will produce a CSV-equivalent
 *
 */
public class ExcelReader
{
    private final ExcelSheetReader excelSheetToCsvConverter;
    private static final String NEW_LINE = System.lineSeparator();

    private final int sheetIndex;
    private final String sheetName;

    private ExcelReader(Builder builder)
    {
        this.sheetIndex = builder.sheetIndex;
        this.sheetName = builder.sheetName;
        this.excelSheetToCsvConverter = new ExcelSheetReader(builder.skipEmptyRows, builder.quoteMode);
    }


    public String[][] createCsvMatrix(String inputFilePath) throws IOException
    {
        return createCsvMatrix(getInputStream(inputFilePath));
    }
    public String[][] createCsvMatrix(InputStream inputStream) throws IOException
    {
        return convertToCsvData(inputStream);
    }

    public String createCsvText(String inputFilePath) throws IOException
    {
        return createCsvText(getInputStream(inputFilePath));
    }
    public String createCsvText(InputStream inputStream) throws IOException
    {
        return convertMatrixToString( convertToCsvData(inputStream) );
    }


    // max method signature permutation limit reached...

    public void createCsvFile(String inputFilePath, String outputFilePath) throws IOException
    {
        createCsvFile(getInputStream(inputFilePath), new File(outputFilePath));
    }
    public void createCsvFile(InputStream inputStream, String outputFilePath) throws IOException
    {
        createCsvFile(inputStream, new File(outputFilePath));
    }
    public void createCsvFile(String inputFilePath, File outputFile) throws IOException
    {
        createCsvFile(getInputStream(inputFilePath), outputFile);
    }
    public void createCsvFile(InputStream inputStream, File outputFile) throws IOException
    {
        String fileText = createCsvText(inputStream);
        FileUtils.writeStringToFile(outputFile, fileText, StandardCharsets.UTF_8);
    }



    private String[][] convertToCsvData(InputStream inputStream) throws IOException
    {
        if (inputStream == null) {
            throw new IllegalArgumentException("Must provide a valid inputStream.");
        }

        String[][] result;
        try (Workbook wb = WorkbookFactory.create(inputStream)) {
            Sheet sheet = getSheet(wb);
            result = excelSheetToCsvConverter.convertToCsvData(sheet);
        }
        finally {
            inputStream.close(); // call close on inputStream b/c Workbook / WorkbookFactory might not.
        }
        return result;
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


    private InputStream getInputStream(String inputFilePath) throws IOException
    {
        if (StringUtils.isEmpty(inputFilePath)) {
            throw new IllegalArgumentException("Must provide a fully-qualified excel input file path");
        }
        if (inputFilePath.startsWith("http") || inputFilePath.startsWith("ftp")) {
            URL url = new URL(inputFilePath);
            final URLConnection connection = url.openConnection();
            connection.setConnectTimeout(20000);
            connection.setReadTimeout(30000);
            return connection.getInputStream();
        }
        else {
            File file;
            if (inputFilePath.startsWith("file:")) {
                URL fileUrl = new URL(inputFilePath);
                file = new File(fileUrl.getFile());
            }
            else {
                file = new File(inputFilePath);
            }

            return new BufferedInputStream(new FileInputStream(file));
        }
    }


    private String convertMatrixToString(String[][] dataMatrix)
    {
        StringBuilder sb = new StringBuilder();

        int columnCount = dataMatrix[0].length;
        int lastColumnIndex = columnCount - 1;

        for (String[] rowData : dataMatrix)
        {
            for (int i = 0; i < columnCount; i++) {
                sb.append(rowData[i]);
                if (i == lastColumnIndex) {
                    sb.append(NEW_LINE);
                }
                else {
                    sb.append(',');
                }
            }
        }
        return sb.toString();
    }




    public static Builder builder() {
        return new Builder();
    }

    public static class Builder {
        private int sheetIndex = 0; // default to the first one.
        private String sheetName = "";  // can optionally provide a specific sheet name
        private boolean skipEmptyRows = false;
        private QuoteMode quoteMode = QuoteMode.NORMAL;

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
