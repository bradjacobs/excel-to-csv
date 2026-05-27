/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.objectmodel;

import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.AbstractExcelReader;
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
public class StandardExcelReader extends AbstractExcelReader {

    private final StandardSheetReader sheetReader;

    // todo: still deciding if this constructor is ok or terrible.
    public StandardExcelReader(SheetConfig sheetConfig) {
        super(sheetConfig);
        this.sheetReader = new StandardSheetReader(sheetConfig);
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
        return sheetInfos;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder extends AbstractSheetConfigBuilder<StandardExcelReader, Builder> {
        @Override
        protected Builder self() {
            return this;
        }

        @Override
        public StandardExcelReader build() {
            return new StandardExcelReader(this.buildConfig());
        }
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
}
