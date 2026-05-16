/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import org.apache.commons.collections4.list.FixedSizeList;
import org.apache.commons.lang3.Validate;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

import static com.github.bradjacobs.excel.util.RowDataUtil.toArray;

/**
 * Represents a mutable implementation of the {@link SheetContent} interface,
 * providing functionalities to manipulate sheet data such as rows, columns,
 * and cell values.
 */
// TODO - class is functional, but still in 'interim' state
public class MutableSheetContent implements SheetContent {

    private static final String EMPTY_VALUE = "";

    private String sheetName;
    private final List<List<String>> rowContent;
    private int columnWidth;

    public static MutableSheetContent copyOf(SheetContent sheetContent) {
        Validate.isTrue(sheetContent != null, "sheetContent cannot be null");
        return new MutableSheetContent(sheetContent.getSheetName(), sheetContent.getRows());
    }

    private MutableSheetContent(String sheetName, List<List<String>> rowContent) {
        setSheetName(sheetName);
        // generate modifiable copy of input rows
        this.rowContent = generateInternalRowData(rowContent);
        // discover the effective column width, and resize rows to match
        this.columnWidth = discoverEffectiveColumnWidth(this.rowContent);
        resizeRowsToEffectiveColumnWidth(this.columnWidth);
    }

    private static List<List<String>> generateInternalRowData(List<List<String>> inputRows) {
        List<List<String>> rows = new ArrayList<>();
        if (inputRows != null) {
            for (List<String> inputRow : inputRows) {
                rows.add(normalizedRowCopy(inputRow));
            }
        }
        return rows;
    }

    @Override
    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = normalizeStringValue(sheetName);
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
        row.set(columnIndex, normalizeStringValue(value));
    }

    @Override
    public List<String> getRow(int rowIndex) {
        return toFixedSizedRow(internalGetRow(rowIndex));
    }

    private List<String> internalGetRow(int rowIndex) {
        validateRowIndex(rowIndex);
        return rowContent.get(rowIndex);
    }

    public void addRow(List<String> rowValues) {
        rowContent.add(copyInputRow(rowValues));
    }

    public void insertRow(int rowIndex, List<String> rowValues) {
        if (rowIndex == rowContent.size()) {
            addRow(rowValues);
        }
        else {
            validateRowIndex(rowIndex);
            rowContent.add(rowIndex, copyInputRow(rowValues));
        }
    }

    public void replaceRow(int rowIndex, List<String> rowValues) {
        validateRowIndex(rowIndex);
        rowContent.set(rowIndex, copyInputRow(rowValues));
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
        return toFixedSizedRows(rowContent);
    }

    @Override
    public String[][] getMatrix() {
        return toArray(rowContent);
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
        fillerValue = normalizeStringValue(fillerValue);
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

        // todo - potential invalid check.. can columIndexList every be empty?
        if (!columIndexList.isEmpty()) {
            this.columnWidth = this.rowContent.get(0).size();
        }
    }

    private void validateRowIndex(int rowIndex) {
        validateIndex(rowIndex, rowContent.size(), "Row");
    }

    private void validateColumnIndex(int columnIndex) {
        validateIndex(columnIndex, columnWidth, "Column");
    }

    private void validateIndex(int index, int size, String label) {
        if (index < 0 || index >= size) {
            throw new IndexOutOfBoundsException(
                    label + " index out of range: " + index + ", size: " + size
            );
        }
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
        // todo - potential index bug below.
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

    private int discoverEffectiveColumnWidth(List<List<String>> rowContent) {
        return rowContent.stream()
                .mapToInt(this::getEffectiveRowWidth)
                .max()
                .orElse(0);
    }

    private int getEffectiveRowWidth(List<String> row) {
        for (int i = row.size() - 1; i >= 0; i--) {
            if (!EMPTY_VALUE.equals(row.get(i))) {
                return i + 1;
            }
        }
        return 0;
    }

    /**
     * Make an editable copy of the input row, and replace any 'null' values with empty string.
     * @param inputRow input row
     * @return copy of input row
     */
    private static List<String> normalizedRowCopy(List<String> inputRow) {
        List<String> copyRow = (inputRow == null) ? new ArrayList<>() : new ArrayList<>(inputRow);
        // replace any 'nulls' with empty string for consistency.
        copyRow.replaceAll(MutableSheetContent::normalizeStringValue);
        return copyRow;
    }

    /**
     * Creates a copy of the input row, adjusts its width to match the required column width,
     * and updates the column width if necessary. Any 'null' values in the input row are
     * replaced with empty strings.
     *
     * @param rowValues the input row to copy and adjust
     * @return a copy of the input row with its width adjusted to the required column width
     */
    private List<String> copyInputRow(List<String> rowValues) {
        List<String> copyRow = normalizedRowCopy(rowValues);

        if (copyRow.size() != columnWidth) {
            int effectiveRowWidth = getEffectiveRowWidth(copyRow);
            if (effectiveRowWidth > columnWidth) {
                columnWidth = effectiveRowWidth;
                resizeRowsToEffectiveColumnWidth(columnWidth);
            }
            resizeRowWidth(copyRow, columnWidth);
        }
        return copyRow;
    }

    private void resizeRowsToEffectiveColumnWidth(int columnWidth) {
        for (List<String> row : rowContent) {
            resizeRowWidth(row, columnWidth);
        }
    }

    private void resizeRowWidth(List<String> row, int desiredWidth) {
        while (row.size() > desiredWidth) {
            row.remove(row.size() - 1);
        }

        while (row.size() < desiredWidth) {
            row.add(EMPTY_VALUE);
        }
    }

    private List<List<String>> toFixedSizedRows(List<List<String>> inputRows) {
        return inputRows.stream()
                .map(this::toFixedSizedRow)
                .collect(Collectors.collectingAndThen(
                        Collectors.toList(),
                        Collections::unmodifiableList));
    }

    private List<String> toFixedSizedRow(List<String> inputRow) {
        return new FixedSizeRow(inputRow);
    }

    private static class FixedSizeRow extends FixedSizeList<String> {
        private static final long serialVersionUID = 1L;

        protected FixedSizeRow(List<String> list) {
            super(list);
        }

        @Override
        public String set(int index, String object) {
            return super.set(index, normalizeStringValue(object));
        }
    }

    /**
     * Converts null to empty string.
     * @param input input value
     * @return normalized value
     */
    private static String normalizeStringValue(String input) {
        // note: currently does not 'trim'.  TBD if this is desired.
        return input != null ? input : EMPTY_VALUE;
    }
}
