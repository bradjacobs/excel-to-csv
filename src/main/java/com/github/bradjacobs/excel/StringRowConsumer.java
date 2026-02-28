/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.UnaryOperator;

/**
 * StringRowConsumer handles processing of String rows from Excel files.
 * Specifically:
 * - logic with normalizing all rows to be the same length
 * - pruning blank rows (if configured)
 * - pruning blank columns (if configured)
 * - any special case logic regarding nulls
 * - Ensure the last row and last column have some data (i.e. not empty)
 * <p>
 * This consumer does _not_ include:
 * - any logic regarding if rows/columns are 'visible' in Excel
 * - trimming or sanitizing cell values
 */
public class StringRowConsumer implements Consumer<List<String>> {

    public enum BlankRemoval {
        NONE(false, false),
        ROWS(true, false),
        COLUMNS(false, true),
        ROWS_AND_COLUMNS(true, true);

        private final boolean removeBlankRows;
        private final boolean removeBlankColumns;

        BlankRemoval(boolean removeBlankRows, boolean removeBlankColumns) {
            this.removeBlankRows = removeBlankRows;
            this.removeBlankColumns = removeBlankColumns;
        }
    }

    private static final String BLANK = "";

    private final List<List<String>> rows = new ArrayList<>();
    private final boolean removeBlankRows;
    private final boolean removeBlankColumns;

    /**
     * Tracks which columns contain at least one non-blank value.
     * Only used when {@link #removeBlankColumns} is enabled.
     */
    private final List<Boolean> keepColumnsFlags = new ArrayList<>();
    private int keepColumnsCount = 0;

    /**
     * The current maximum column count observed after normalizing rows.
     */
    private int maxColumnCount = 0;


    public static StringRowConsumer of(BlankRemoval removal) {
        if (removal == null) {
            removal = BlankRemoval.NONE;
        }
        return of(removal.removeBlankRows, removal.removeBlankColumns);
    }

    public static StringRowConsumer of(
            boolean removeBlankRows,
            boolean removeBlankColumns) {
        return new StringRowConsumer(removeBlankRows, removeBlankColumns);
    }

    private StringRowConsumer(boolean removeBlankRows, boolean removeBlankColumns) {
        this.removeBlankRows = removeBlankRows;
        this.removeBlankColumns = removeBlankColumns;
    }

    /**
     * Consumes all the String value rows read from the Excel file.
     */
    @Override
    public void accept(List<String> row) {
        if (removeBlankRows && isEmptyRow(row)) {
            return;
        }
        rows.add(normalizeRow(row));
    }


    public List<List<String>> generateMatrixList() {
        removeTrailingBlankRows();

        // Ensure all rows have uniform width before potentially filtering columns.
        for (List<String> row : rows) {
            padRowRightToWidth(row, maxColumnCount);
        }

        filterBlankColumnsIfNeeded();
        return Collections.unmodifiableList(rows);
    }

    public String[][] generateMatrix() {
        List<List<String>> matrixList = generateMatrixList();
        return matrixList.stream()
                .map(inner -> inner.toArray(new String[0]))
                .toArray(String[][]::new);
    }

    private List<String> normalizeRow(List<String> input) {
        // wrap the input list with our own ArrayList, since
        // the input list might be a read-only list.
        List<String> row = (input == null) ? new ArrayList<>() : new ArrayList<>(input);

        // replace any 'nulls' with empty string for consistency.
        row.replaceAll(s -> s == null ? BLANK : s);

        // If we previously saw a wider row, pad this one.
        padRowRightToWidth(row, maxColumnCount);

        // If this row is wider than the current max, trim trailing blanks (if any),
        // then update the max column count and (optionally) the column bookkeeping structures.
        if (row.size() > maxColumnCount) {
            trimRowTrailingBlanksPastMaxWidth(row, maxColumnCount);
            maxColumnCount = row.size();
            ensureColumnTrackingCapacity(maxColumnCount);
        }

        updateKeepColumnFlags(row);
        return row;
    }

    private static void padRowRightToWidth(List<String> row, int targetWidth) {
        while (row.size() < targetWidth) {
            row.add(BLANK);
        }
    }

    /**
     * Removes trailing blank cells beyond the previous max width.
     * Keeps data if a non-blank value is encountered.
     */
    private static void trimRowTrailingBlanksPastMaxWidth(List<String> row, int previousMaxWidth) {
        for (int i = row.size() - 1; i >= previousMaxWidth; i--) {
            String value = row.get(i);
            if (StringUtils.isNotEmpty(value))  {
                break;
            }
            row.remove(i);
        }
    }

    private void ensureColumnTrackingCapacity(int width) {
        if (!removeBlankColumns) {
            return;
        }
        while (keepColumnsFlags.size() < width) {
            keepColumnsFlags.add(Boolean.FALSE);
        }
    }

    private void updateKeepColumnFlags(List<String> row) {
        // bail early if not configured to remove blank columns.
        if (!removeBlankColumns) {
            return;
        }
        // if all tracked columns have already been detected
        // to have a value, then return.
        if (keepColumnsCount >= keepColumnsFlags.size()) {
            return;
        }

        int limit = Math.min(row.size(), keepColumnsFlags.size());
        for (int i = 0; i < limit; i++) {
            // if the column flags indicate there is no value
            // for the current column and the given row does have
            // a value for the column, then set the keepColumnFlag to true.
            if (!keepColumnsFlags.get(i) && !row.get(i).isEmpty()) {
                keepColumnsFlags.set(i, Boolean.TRUE);
                keepColumnsCount++;
                if (keepColumnsCount >= keepColumnsFlags.size()) {
                    return;
                }
            }
        }
    }

    private void filterBlankColumnsIfNeeded() {
        if (!shouldFilterBlankColumns()) {
            return;
        }
        rows.replaceAll(new FilterColumnsOperator(keepColumnsFlags, keepColumnsCount));
    }

    /**
     * Simple check to see if any blank column removal is needed.
     * @return true if there are one or more blank columns to remove.
     */
    private boolean shouldFilterBlankColumns() {
        return removeBlankColumns
                && !rows.isEmpty()
                && keepColumnsCount < keepColumnsFlags.size();
    }

    // Helper class to remove blank column(s) from each row
    private static class FilterColumnsOperator implements UnaryOperator<List<String>> {
        private final List<Boolean> keepColumnsFlags;
        private final int keepColumnsCount;

        public FilterColumnsOperator(List<Boolean> keepColumnsFlags, int keepColumnsCount) {
            this.keepColumnsFlags = keepColumnsFlags;
            this.keepColumnsCount = keepColumnsCount;
        }

        @Override
        public List<String> apply(List<String> row) {
            List<String> filtered = new ArrayList<>(keepColumnsCount);
            int limit = Math.min(row.size(), keepColumnsFlags.size());
            for (int i = 0; i < limit; i++) {
                if (Boolean.TRUE.equals(keepColumnsFlags.get(i))) {
                    filtered.add(row.get(i));
                }
            }
            return filtered;
        }
    }

    /**
     * Trim off any trailing blank rows (so the last row always has data)
     */
    private void removeTrailingBlankRows() {
        while (!rows.isEmpty() && isEmptyRow(rows.get(rows.size() - 1))) {
            rows.remove(rows.size() - 1);
        }
    }

    /**
     * Check if row is considered to be 'blank/empty'
     */
    private static boolean isEmptyRow(List<String> rowData) {
        if (CollectionUtils.isEmpty(rowData)) {
            return true;
        }
        for (String value : rowData) {
            if (StringUtils.isNotEmpty(value)) {
                return false;
            }
        }
        return true;
    }
}