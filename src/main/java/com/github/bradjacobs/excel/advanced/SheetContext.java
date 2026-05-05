/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import java.util.HashSet;
import java.util.Set;

/**
 * Tracks hidden row/column indices while parsing a sheet.
 */
final class SheetContext {
    private final Set<Integer> hiddenRows = new HashSet<>();
    private final Set<Integer> hiddenColumns = new HashSet<>();

    public boolean isRowVisible(int rowIndex) {
        return !hiddenRows.contains(rowIndex);
    }

    public boolean isColumnVisible(int columnIndex) {
        return !hiddenColumns.contains(columnIndex);
    }

    public void addHiddenRow(int rowIndex) {
        hiddenRows.add(rowIndex);
    }

    public void addHiddenColumn(int columnIndex) {
        hiddenColumns.add(columnIndex);
    }
}
