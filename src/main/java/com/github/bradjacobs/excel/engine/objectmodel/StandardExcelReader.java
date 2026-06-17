/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.objectmodel;

import com.github.bradjacobs.excel.api.ExcelWorkbookReader;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.model.SheetContent;
import com.github.bradjacobs.excel.request.ExcelReadRequest;
import com.github.bradjacobs.excel.request.SheetInfo;
import com.github.bradjacobs.excel.request.SheetSelector;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Reads an Excel File sheets
 */
public class StandardExcelReader implements ExcelWorkbookReader {

    private final StandardSheetReader sheetReader;

    private StandardExcelReader(SheetConfig config) {
        this.sheetReader = new StandardSheetReader(config);

        // override the internal POI utils size limit to allow for 'bigger Excel files'
        //   (as of POI version 5.2.0 the default value is 100_000_000)
        org.apache.poi.util.IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
    }

    @Override
    public List<SheetContent> readSheets(ExcelReadRequest request) throws IOException {
        Validate.isTrue(request != null, "Request cannot be null");

        String password = request.getPassword();
        SheetSelector sheetSelector = request.getSheetSelector();
        InputStream excelInputStream = request.getSourceInputStream();

        try (excelInputStream; Workbook workbook = WorkbookFactory.create(excelInputStream, password)) {
            List<WorkbookSheetInfo> allSheets = getWorkbookSheets(workbook);
            List<WorkbookSheetInfo> selectedSheets = sheetSelector.filterSheets(allSheets);
            return toSheetContents(selectedSheets);
        }
    }

    private List<SheetContent> toSheetContents(List<WorkbookSheetInfo> selectedSheets) {
        return selectedSheets.stream()
                .map(wsi -> sheetReader.toSheetContent(wsi.getSheet()))
                .collect(Collectors.toList());
    }

    /**
     * Gets all sheets in the given workbook
     * returns a list of 'WorkbookSheetInfo', which includes: Sheet, SheetName, SheetIndex.
     * @param workbook workbook
     * @return list of workbook sheet infos (which contain the sheet object)
     */
    private List<WorkbookSheetInfo> getWorkbookSheets(Workbook workbook) {
        List<WorkbookSheetInfo> sheetInfos = new ArrayList<>();
        int sheetIndex = 0;
        for (Sheet sheet : workbook) {
            sheetInfos.add(new WorkbookSheetInfo(sheet.getSheetName(), sheetIndex++, sheet));
        }
        if (sheetInfos.isEmpty()) {
            throw new IllegalStateException("Unable to read workbook sheets. File might be corrupted.");
        }
        return sheetInfos;
    }

    private static class WorkbookSheetInfo implements SheetInfo {
        private final String sheetName;
        private final int sheetIndex;
        private final Sheet sheet;

        public WorkbookSheetInfo(String sheetName, int sheetIndex, Sheet sheet) {
            this.sheetName = sheetName;
            this.sheetIndex = sheetIndex;
            this.sheet = sheet;
        }

        @Override
        public String getName() {
            return sheetName;
        }

        @Override
        public int getIndex() {
            return sheetIndex;
        }

        public Sheet getSheet() {
            return sheet;
        }
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder extends SheetConfig.AbstractSheetConfigBuilder<StandardExcelReader, Builder> {
        @Override
        protected Builder self() {
            return this;
        }

        @Override
        public StandardExcelReader build() {
            return new StandardExcelReader(this.buildConfig());
        }
    }
}
