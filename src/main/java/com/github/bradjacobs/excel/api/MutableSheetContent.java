/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import org.apache.commons.collections4.list.FixedSizeList;
import org.apache.commons.lang3.Validate;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * Represents a mutable implementation of the {@link SheetContent} interface,
 * providing functionalities to manipulate sheet data such as rows, columns,
 * and cell values.
 */
public class MutableSheetContent implements SheetContent {

    private static final String EMPTY_VALUE = "";

    private String sheetName;
    private final List<List<String>> rowContent;
    private int columnWidth;

    public static MutableSheetContent from(SheetContent sheetContent) {
        // TODO refactor
        String[][] matrix = sheetContent.getMatrix();
        List<List<String>> listData = new ArrayList<>(matrix.length);
        for (String[] inner : matrix) {
            listData.add(new ArrayList<>(Arrays.asList(inner)));
        }
        return new MutableSheetContent(sheetContent.getSheetName(), listData);
    }

    private MutableSheetContent(String sheetName, List<List<String>> rowContent) {
        setSheetName(sheetName);
        this.rowContent = rowContent;
        this.columnWidth = rowContent.isEmpty() ? 0 : rowContent.get(0).size();
    }

    @Override
    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName != null ? sheetName : EMPTY_VALUE;
    }

    @Override
    public boolean isEmpty() {
        return rowContent.isEmpty();
    }

    @Override
    public int getRowCount() {
        return rowContent.size();
    }

    @Override
    public int getColumnCount() {
        return columnWidth;
    }

    @Override
    public String getCellValue(int rowIndex, int columnIndex) {
        List<String> row = internalGetRow(rowIndex);
        validateColumnIndex(columnIndex);
        return row.get(columnIndex);
    }

    public void setCellValue(int rowIndex, int columnIndex, String value) {
        List<String> row = internalGetRow(rowIndex);
        validateColumnIndex(columnIndex);
        row.set(columnIndex, value);
    }

    @Override
    public List<String> getRow(int rowIndex) {
        //return internalGetRow(rowIndex);
        return new FixedSizeRow(internalGetRow(rowIndex));
        //return copyRow(internalGetRow(rowIndex));
    }

    private List<String> internalGetRow(int rowIndex) {
        validateRowIndex(rowIndex);
        return rowContent.get(rowIndex);
    }

    public void addRow(List<String> rowValues) {
        rowContent.add(copyRow(rowValues));
    }

    public void insertRow(int rowIndex, List<String> rowValues) {
        if (rowIndex == rowContent.size()) {
            addRow(rowValues);
        }
        else {
            validateRowIndex(rowIndex);
            rowContent.add(rowIndex, copyRow(rowValues));
        }
    }

    public void replaceRow(int rowIndex, List<String> rowValues) {
        validateRowIndex(rowIndex);
        rowContent.set(rowIndex, copyRow(rowValues));
    }

    public void transformRow(int rowIndex, Function<String, String> transformer) {
        Validate.isTrue(transformer != null,
                "transformer must not be null");
        List<String> row = internalGetRow(rowIndex);
        row.replaceAll(transformer::apply);
    }

    public List<String> removeRow(int rowIndex) {
        validateRowIndex(rowIndex);
        return rowContent.remove(rowIndex);
    }

    @Override
    public List<List<String>> getRows() {
        return List.copyOf(
                rowContent.stream()
                        .map(List::copyOf)
                        .collect(Collectors.toList()));
    }

    @Override
    public String[][] getMatrix() {
        return rowContent.stream()
                .map(inner -> inner.toArray(new String[0]))
                .toArray(String[][]::new);
    }

    public void transformCells(Function<String, String> transformer) {
        Validate.isTrue(transformer != null,
                "transformer must not be null");

        for (List<String> currentRow : rowContent) {
            currentRow.replaceAll(transformer::apply);
        }
    }

    public void addColumn() {
        addColumn(EMPTY_VALUE);
    }

    public void addColumn(String fillerValue) {
        fillerValue = fillerValue != null ? fillerValue : EMPTY_VALUE;
        for (List<String> row : rowContent) {
            row.add(fillerValue);
        }
        columnWidth++;
    }

    public void removeColumn(int... columnIndexes) {
        if (columnIndexes == null || columnIndexes.length == 0) {
            return;
        }

        List<Integer> columIndexList = Arrays.stream(columnIndexes)
                .distinct()
                .boxed()
                .sorted(Comparator.reverseOrder()) // important to be in 'reverse order'!
                .collect(Collectors.toList());

        // first ensure all indexes are valid.
        for (Integer columnIndex : columIndexList) {
            validateColumnIndex(columnIndex);
        }

        for (Integer columnIndex : columIndexList) {
            for (List<String> row : rowContent) {
                row.remove((int)columnIndex);
            }
        }

        if (!columIndexList.isEmpty()) {
            this.columnWidth = this.rowContent.get(0).size();
        }
    }

    private void validateRowIndex(int rowIndex) {
        if (rowIndex < 0 || rowIndex >= rowContent.size()) {
            throw new IndexOutOfBoundsException(
                    "Row index out of range: " + rowIndex + ", size: " +  rowContent.size()
            );
        }
    }

    private void validateColumnIndex(int columnIndex) {
        if (columnIndex < 0 || columnIndex >= columnWidth) {
            throw new IndexOutOfBoundsException(
                    "Column index out of range: " + columnIndex + ", size: " +  columnWidth
            );
        }
    }

    // todo rename method
    private List<String> copyRow(List<String> rowValues) {
        if (rowValues == null) {
            rowValues = List.of();
        }

        Validate.isTrue(rowValues.size() <= columnWidth,
                "Row contains too many columns: " + rowValues.size() + " > " + columnWidth);
        List<String> copyRow = rowValues.stream()
                .map(s -> (s == null) ? "" : s)
                .collect(Collectors.toCollection(ArrayList::new));

        while (copyRow.size() < columnWidth) {
            copyRow.add(EMPTY_VALUE);
        }
        return copyRow;
    }

    // todo - behavior if pass in a header that doesn't exist?
    //   probably just ignore.
    public void removeColumn(String... headerNames) {
        if (headerNames == null || headerNames.length == 0) {
            return;
        }

        Set<String> headersToRemove = toHeaderSet(headerNames);
        removeColumn(findColumnIndexesByHeaderNames(headersToRemove));
    }

    private Set<String> toHeaderSet(String... headerNames) {
        return Arrays.stream(headerNames)
                .filter(Objects::nonNull)
                .collect(Collectors.toSet());
    }

    private int[] findColumnIndexesByHeaderNames(Set<String> headerNames) {
        List<String> headerRow = internalGetRow(0);
        List<Integer> columnIndexes = new ArrayList<>();

        for (int columnIndex = 0; columnIndex < headerRow.size(); columnIndex++) {
            if (headerNames.contains(headerRow.get(columnIndex))) {
                columnIndexes.add(columnIndex);
            }
        }

        return columnIndexes.stream()
                .mapToInt(Integer::intValue)
                .toArray();
    }


    private static class FixedSizeRow extends FixedSizeList<String> {
        protected FixedSizeRow(List<String> list) {
            super(list);
        }
    }
}
