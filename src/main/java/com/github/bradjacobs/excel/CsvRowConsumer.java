/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.function.Consumer;

// TODO - merge this class into the ExcelSheetReader class.
public class CsvRowConsumer implements Consumer<String[]> {
    private final List<String[]> rowList = new ArrayList<>();
    private final boolean removeEmptyRows;
    private final boolean removeEmptyColumns;
    private boolean[] columnsWithDataArray = null;
    private int numberOfColumnsWithData = 0;

    public CsvRowConsumer(boolean removeEmptyRows, boolean removeEmptyColumns) {
        this.removeEmptyRows = removeEmptyRows;
        this.removeEmptyColumns = removeEmptyColumns;
    }

    @Override
    public void accept(String[] row) {
        // Remove blank rows if configured
        if (this.removeEmptyRows && isEmptyRow(row))
            return;

        if (this.removeEmptyColumns) {
            if (columnsWithDataArray == null) {
                columnsWithDataArray = new boolean[row.length];
            }
            if (numberOfColumnsWithData < columnsWithDataArray.length) {
                for (int i = 0; i < columnsWithDataArray.length; i++) {
                    if (!columnsWithDataArray[i] && !row[i].isEmpty()) {
                        columnsWithDataArray[i] = true;
                        numberOfColumnsWithData++;
                    }
                }
            }
        }
        this.rowList.add(row);
    }

    public List<String[]> getRows() {
        return Collections.unmodifiableList(rowList);
    }

    public void finalizeRows() {
        // remove any trailing blank rows (regardless of 'skipEmptyRows' value)
        while (!rowList.isEmpty() && isEmptyRow(rowList.get(rowList.size()-1))) {
            rowList.remove(rowList.size()-1);
        }

        if (removeEmptyColumns &&
                !rowList.isEmpty() &&
                columnsWithDataArray != null &&
                numberOfColumnsWithData < columnsWithDataArray.length) {

            // todo: this is not good if excel sheet is super big
            //   because would end up having 2 copies in memory.
            List<String[]> filteredRows = new ArrayList<>();
            for (String[] row : rowList) {
                filteredRows.add(filterColumns(row, columnsWithDataArray, numberOfColumnsWithData));
            }
            rowList.clear();
            rowList.addAll(filteredRows);
        }
    }

    private String[] filterColumns(String[] row, boolean[] columnHasData, int count) {
        String[] filtered = new String[count];
        for (int i = 0, idx = 0; i < columnHasData.length && i < row.length; i++) {
            if (columnHasData[i]) {
                filtered[idx++] = row[i];
            }
        }
        return filtered;
    }

    protected boolean isEmptyRow(String[] rowData) {
        for (String r : rowData) {
            if (r != null && !r.isEmpty()) {
                return false;
            }
        }
        return true;
    }
}
