/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.row;

import org.apache.commons.collections4.list.UnmodifiableList;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Utility methods for converting row-oriented spreadsheet data between arrays and lists.
 * <p>
 * This class also provides helpers for wrapping rows and collections of rows in
 * unmodifiable list views, while treating {@code null} inputs as empty results.
 */
public class RowDataUtil {

    private static final List<String> EMPTY_ROW = new UnmodifiableList<>(List.of());
    private static final List<List<String>> EMPTY_ROWS = new UnmodifiableList<>(List.of());

    public static List<List<String>> toUnmodifiableRows(String[][] array) {
        if (array == null) {
            return EMPTY_ROWS;
        }
        List<List<String>> readOnlyRows = new ArrayList<>(array.length);
        for (String[] row : array) {
            readOnlyRows.add(toUnmodifiableRow(row));
        }
        return new UnmodifiableList<>(readOnlyRows);
    }

    public static List<List<String>> toUnmodifiableRows(List<List<String>> rows) {
        if (rows == null) {
            return EMPTY_ROWS;
        }
        else if (rows instanceof UnmodifiableList) {
            return rows;
        }
        List<List<String>> readOnlyRows = new ArrayList<>(rows.size());
        for (List<String> row : rows) {
            readOnlyRows.add(toUnmodifiableRow(row));
        }
        return new UnmodifiableList<>(readOnlyRows);
    }

    public static List<String> toUnmodifiableRow(String[] row) {
        if (row == null) {
            return EMPTY_ROW;
        }
        return toUnmodifiableRow(Arrays.asList(row));
    }

    public static List<String> toUnmodifiableRow(List<String> row) {
        if (row == null) {
            return EMPTY_ROW;
        }
        else if (row instanceof UnmodifiableList) {
            return row;
        }
        return new UnmodifiableList<>(row);
    }

    public static String[][] toArray(List<List<String>> rows) {
        if (rows == null) {
            return new String[0][0];
        }
        int rowCount = rows.size();
        String[][] result = new String[rowCount][];
        for (int i = 0; i <rowCount; i++) {
            List<String> row = rows.get(i);
            result[i] = row.toArray(new String[0]);
        }
        return result;
    }
}
