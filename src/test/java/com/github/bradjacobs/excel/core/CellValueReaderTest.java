/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import com.github.bradjacobs.excel.config.SanitizeType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.jupiter.api.Test;

import java.util.Set;

import static com.github.bradjacobs.excel.config.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.config.SanitizeType.QUOTES;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

/**
 * Basic UnitTests for the CellValueReader.
 */
// NOTE: limited tests here because don't want to learn
//   about mocking all the different internal variations of cell types.
class CellValueReaderTest {

    private static final String CELL_VALUE_MISMATCH_MESSAGE = "mismatch expected cell value";
    private static final String NULL_SANITIZE_TYPES_MESSAGE = "Must provide non-null sanitizeTypes.";

    @Test
    public void trimsStringCellValueWhenTrimEnabled() {
        String inputString = "  dog  ";
        String expectedString = inputString.trim();
        CellValueReader cellValueReader = createCellValueReader(true, Set.of());
        Cell cell = createMockStringCell(inputString);
        String result = cellValueReader.getCellValue(cell);
        assertEquals(expectedString, result, CELL_VALUE_MISMATCH_MESSAGE);
    }

    @Test
    public void preservesStringCellValueWhenTrimDisabled() {
        String inputString = "  dog  ";
        CellValueReader cellValueReader = createCellValueReader(false, Set.of());
        Cell cell = createMockStringCell(inputString);
        String result = cellValueReader.getCellValue(cell);
        assertEquals(inputString, result, CELL_VALUE_MISMATCH_MESSAGE);
    }

    @Test
    public void removesBasicDiacriticsWhenConfigured() {
        CellValueReader cellValueReader = createCellValueReader(true, Set.of(BASIC_DIACRITICS));
        Cell cell = createMockStringCell("Façade");
        String result = cellValueReader.getCellValue(cell);
        assertEquals("Facade", result, CELL_VALUE_MISMATCH_MESSAGE);
    }

    @Test
    public void returnsEmptyStringForNullCell() {
        CellValueReader cellValueReader = createCellValueReader(true, Set.of(QUOTES));
        String result = cellValueReader.getCellValue(null);
        assertEquals("", result, CELL_VALUE_MISMATCH_MESSAGE);
    }

    @Test
    public void readsCachedNumericValueFromFormulaCell() {
        // note: this is a semi-poor representation of a true Excel Formula cell.
        CellValueReader cellValueReader = createCellValueReader(true, Set.of(QUOTES));
        Cell cell = mock(Cell.class);
        when(cell.getCellType()).thenReturn(CellType.FORMULA);
        when(cell.getCachedFormulaResultType()).thenReturn(CellType.NUMERIC);
        when(cell.getCellFormula()).thenReturn("A1+B1"); // cellFormula string is only used if the code is misconfigured
        when(cell.getNumericCellValue()).thenReturn(31.2d);

        String result = cellValueReader.getCellValue(cell);

        assertEquals("31.2", result, CELL_VALUE_MISMATCH_MESSAGE);
    }

    @Test
    public void throwsExceptionWhenSanitizeTypesIsNull() {
        Exception exception = assertThrows(
                IllegalArgumentException.class,
                () -> new CellValueReader(true, null)
        );
        assertEquals(NULL_SANITIZE_TYPES_MESSAGE, exception.getMessage(), "Mismatch exception message");
    }

    private CellValueReader createCellValueReader(boolean trimStringValues, Set<SanitizeType> sanitizeTypes) {
        return new CellValueReader(trimStringValues, sanitizeTypes);
    }

    /**
     * Helper method for creating a Mock Cell input
     *
     * @param cellValue string value for cell
     * @return cell
     */
    private Cell createMockStringCell(String cellValue) {
        Cell cell = mock(Cell.class);
        when(cell.getCellType()).thenReturn(CellType.STRING);
        when(cell.getRichStringCellValue()).thenReturn(new XSSFRichTextString(cellValue));
        return cell;
    }
}
