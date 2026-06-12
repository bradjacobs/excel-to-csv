/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.shared;

import java.util.HashSet;
import java.util.Set;

/**
 * Tracks hidden row/column indices while parsing a sheet.
 */
public final class SheetVisibilityTracker implements SheetContentHandler.SheetContentEmitPolicy {
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

    @Override
    public boolean shouldEmitRow(int rowIndex) {
        return isRowVisible(rowIndex);
    }

    @Override
    public boolean shouldEmitColumn(int columnIndex) {
        return isColumnVisible(columnIndex);
    }
}
