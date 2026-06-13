/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.shared;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Simple test of {@link SheetVisibilityTracker} that handles
 * tracking visibility of rows and columns.
 * Any value not explicitly marked as hidden is considered visible.
 */
class SheetVisibilityTrackerTest {

    @Test
    void rowsAreVisibleByDefault() {
        SheetVisibilityTracker tracker = new SheetVisibilityTracker();
        assertTrue(tracker.isRowVisible(0));
        assertTrue(tracker.isRowVisible(10));
    }

    @Test
    void columnsAreVisibleByDefault() {
        SheetVisibilityTracker tracker = new SheetVisibilityTracker();
        assertTrue(tracker.isColumnVisible(0));
        assertTrue(tracker.isColumnVisible(5));
    }

    @Test
    void hiddenRowBecomesInvisible() {
        SheetVisibilityTracker tracker = new SheetVisibilityTracker();
        tracker.addHiddenRow(3);
        assertFalse(tracker.isRowVisible(3));
        assertTrue(tracker.isRowVisible(4));
    }

    @Test
    void hiddenColumnBecomesInvisible() {
        SheetVisibilityTracker tracker = new SheetVisibilityTracker();
        tracker.addHiddenColumn(7);
        assertFalse(tracker.isColumnVisible(7));
        assertTrue(tracker.isColumnVisible(8));
    }

    @Test
    void shouldEmitRowReturnsTrueForVisibleRow() {
        SheetVisibilityTracker tracker = new SheetVisibilityTracker();
        assertTrue(tracker.shouldEmitRow(1));
    }

    @Test
    void shouldEmitRowReturnsFalseForHiddenRow() {
        SheetVisibilityTracker tracker = new SheetVisibilityTracker();
        tracker.addHiddenRow(1);
        assertFalse(tracker.shouldEmitRow(1));
    }

    @Test
    void shouldEmitColumnReturnsTrueForVisibleColumn() {
        SheetVisibilityTracker tracker = new SheetVisibilityTracker();
        assertTrue(tracker.shouldEmitColumn(2));
    }

    @Test
    void shouldEmitColumnReturnsFalseForHiddenColumn() {
        SheetVisibilityTracker tracker = new SheetVisibilityTracker();
        tracker.addHiddenColumn(2);
        assertFalse(tracker.shouldEmitColumn(2));
    }
}
