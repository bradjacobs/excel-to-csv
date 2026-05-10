package com.github.bradjacobs.excel.util;

import org.apache.commons.collections4.list.UnmodifiableList;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class RowDataUtil {

    public static List<List<String>> toUnmodifiableRows(String[][] array) {
        if (array == null) {
            return List.of();
        }
        List<List<String>> readOnlyRows = new ArrayList<>(array.length);
        for (String[] row : array) {
            readOnlyRows.add(toUnmodifiableRow(row));
        }
        return new UnmodifiableList<>(readOnlyRows);
    }

    public static List<List<String>> toUnmodifiableRows(List<List<String>> rows) {
        if (rows instanceof UnmodifiableList) {
            return rows;
        }
        List<List<String>> readOnlyRows = new ArrayList<>(rows.size());
        for (List<String> row : rows) {
            readOnlyRows.add(toUnmodifiableRow(row));
        }
        return new UnmodifiableList<>(readOnlyRows);
    }

    public static List<String> toUnmodifiableRow(String[] row) {
        return toUnmodifiableRow(Arrays.asList(row));
    }

    public static List<String> toUnmodifiableRow(List<String> row) {
        return new UnmodifiableList<>(row);
    }

    public static String[][] toArray(List<List<String>> rows) {
        int rowCount = rows.size();
        String[][] result = new String[rowCount][];
        for (int i = 0; i <rowCount; i++) {
            List<String> row = rows.get(i);
            result[i] = row.toArray(new String[0]);
        }
        return result;
    }

}
