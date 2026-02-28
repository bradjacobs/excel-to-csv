/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;

import java.util.Arrays;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;

public class StringRowConsumerTest {

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BasicConsumerTests {
        @Test
        public void basicRowAccept() {
            String[][] input = {{"value1", "value2"}};
            String[][] expected = {{"value1", "value2"}};
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void noInputRowAccept() {
            String[][] input = {};
            String[][] expected = {};
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void consumeNullRow() {
            String[][] input = {
                    {"aa", "aa"},
                    null,
                    {"cc", "cc"}
            };
            String[][] expected = {
                    {"aa", "aa"},
                    {"", ""},
                    {"cc", "cc"}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void consumeEmptyRow() {
            String[][] input = {
                    {"aa", "aa"},
                    {},
                    {"cc", "cc"}
            };
            String[][] expected = {
                    {"aa", "aa"},
                    {"", ""},
                    {"cc", "cc"}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void consumeEmptyFirstRow() {
            String[][] input = {
                    {},
                    {"bb", "bb"}
            };
            String[][] expected = {
                    {"", ""},
                    {"bb", "bb"}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void conumeWithNullInLastColumn() {
            String[][] input = {
                    {"aa", null},
                    {},
                    {"cc", "cc"}
            };
            String[][] expected = {
                    {"aa", ""},
                    {"", ""},
                    {"cc", "cc"}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void consumeRowWithNullValue() {
            String[][] input = {
                    {"aa", null, "bb"}
            };
            String[][] expected = {
                    {"aa", "", "bb"}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }


        @Test
        public void longestRowFirst() {
            String[][] input = {
                    {"aa", "aa", "aa", "aa"},
                    {"bb", "bb"},
                    {"cc", "cc", "cc"}
            };
            String[][] expected = {
                    {"aa", "aa", "aa", "aa"},
                    {"bb", "bb", "", ""},
                    {"cc", "cc", "cc", ""}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void longestRowMiddle() {
            String[][] input = {
                    {"aa", "aa"},
                    {"bb", "", "bb", "bb"},
                    {"cc"}
            };
            String[][] expected = {
                    {"aa", "aa", "", ""},
                    {"bb", "", "bb", "bb"},
                    {"cc", "", "", ""}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void longestRowLast() {
            String[][] input = {
                    {"aa"},
                    {"bb", "bb"},
                    {"cc", "cc", "cc"}
            };
            String[][] expected = {
                    {"aa", "", ""},
                    {"bb", "bb", ""},
                    {"cc", "cc", "cc"}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void fluctuatingRowSizes() {
            String[][] input = {
                    {"aa", "aa", "aa"},
                    {"bb"},
                    {"cc", "cc", "cc", "cc", "cc"},
                    {"dd", "dd"}
            };
            String[][] expected = {
                    {"aa", "aa", "aa", "", ""},
                    {"bb", "", "", "", ""},
                    {"cc", "cc", "cc", "cc", "cc"},
                    {"dd", "dd", "", "", ""}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void longestBlankRow() {
            String[][] input = {
                    {"aa", "aa"},
                    {"", "", "", "", ""},
                    {"cc"}
            };
            String[][] expected = {
                    {"aa", "aa"},
                    {"", ""},
                    {"cc", ""}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        @Test
        public void startingBlankColumns() {
            String[][] input = {
                    {"", "aa"},
                    {"", "", "bb"},
                    {"", "cc"}
            };
            String[][] expected = {
                    {"", "aa", ""},
                    {"", "", "bb"},
                    {"", "cc", ""}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }


        // beginning blanks reows are to be kept,
        //   if not configured to remove blank rows.
        @Test
        public void retainFirstlankRows() {
            String[][] input = {
                    {"", ""},
                    {"", ""},
                    {"aa", "bb"},
                    {"cc", "dd"},
            };
            String[][] expected = {
                    {"", ""},
                    {"", ""},
                    {"aa", "bb"},
                    {"cc", "dd"},
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        // prune all final blank rows, regardless of configuration.
        @Test
        public void pruneTrailingBlankRows() {
            String[][] input = {
                    {"bb", "bb"},
                    {"", ""},
                    {"dd", "dd"},
                    {"", ""},
                    {"", ""},
            };
            String[][] expected = {
                    {"bb", "bb"},
                    {"", ""},
                    {"dd", "dd"}
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }

        // prune all final blank columns, regardless of configuration.
        @Test
        public void pruneTrailingBlankColumns() {
            String[][] input = {
                    {"aa", "bb", "", ""},
                    {"cc"},
                    {"dd", "ee", ""},
                    {"ff", "gg"},
            };
            String[][] expected = {
                    {"aa", "bb"},
                    {"cc", ""},
                    {"dd", "ee"},
                    {"ff", "gg"},
            };
            runConsumerTest(createBasicConsumer(), input, expected);
        }


        @Test
        public void createWithNullFactoryParamIsNoneBehavior() {
            // if you create a consumer with a 'null'
            // should behave to keep blank rows and columsn
            String[][] input = {
                    {"aa", "", "bb"},
                    {"", "", ""},
                    {"cc", "", "dd"}
            };
            String[][] expected = {
                    {"aa", "", "bb"},
                    {"", "", ""},
                    {"cc", "", "dd"}
            };
            StringRowConsumer consumer = StringRowConsumer.of(null);
            runConsumerTest(consumer, input, expected);
        }



    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class PruneEmptyRowConsumerTests {
        @Test
        public void removesBlankRowsWhenConfigured() {
            String[][] input = {
                    {"aa", "bb"},
                    {"", ""},
                    {"cc", "dd"}
            };
            String[][] expected = {
                    {"aa", "bb"},
                    {"cc", "dd"}
            };
            runConsumerTest(createPruneBlankRowsConsumer(), input, expected);
        }

        @Test
        public void removeMultiBlankRowsWhenConfigured() {
            String[][] input = {
                    {"", "", ""},
                    {"", ""},
                    {"aa", "bb"},
                    null,
                    {"", ""},
                    {"", ""},
                    {"cc", "dd"},
                    {"", ""},
            };
            String[][] expected = {
                    {"aa", "bb"},
                    {"cc", "dd"}
            };
            runConsumerTest(createPruneBlankRowsConsumer(), input, expected);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class PruneEmptyColumnConsumerTests {
        @Test
        public void basicPruneBlankColumns() {
            String[][] input = {
                    {"aa", "", "cc"},
                    {"dd", "", "ff"},
            };
            String[][] expected = {
                    {"aa", "cc"},
                    {"dd", "ff"},
            };
            runConsumerTest(createPruneBlankColumnConsumer(), input, expected);
        }

        @Test
        public void basicBlankColumnsDifferentRowLengths() {
            String[][] input = {
                    {"", "", "cc"},
                    {"dd", "", "ff", "", "hh"},
                    {"rr", ""},
                    {"ss", "", "tt", "", "vv", ""},
            };
            String[][] expected = {
                    {"", "cc", ""},
                    {"dd", "ff", "hh"},
                    {"rr", "", ""},
                    {"ss", "tt", "vv"},
            };
            runConsumerTest(createPruneBlankColumnConsumer(), input, expected);
        }

        @Test
        public void pruneStartingBlankColumns() {
            String[][] input = {
                    {"", "", "cc"},
                    {"", "", "dd"},
                    {"", "", "ee", "ff"},
                    {"", "", "hh"},
            };
            String[][] expected = {
                    {"cc", ""},
                    {"dd", ""},
                    {"ee", "ff"},
                    {"hh", ""},
            };
            runConsumerTest(createPruneBlankColumnConsumer(), input, expected);
        }

        @Test
        public void pruneEndingBlankColumns() {
            String[][] input = {
                    {"cc", ""},
                    {"dd", "", ""},
                    {"ee", ""},
                    {"ff", ""},
            };
            String[][] expected = {
                    {"cc"},
                    {"dd"},
                    {"ee"},
                    {"ff"},
            };
            runConsumerTest(createPruneBlankColumnConsumer(), input, expected);
        }

        @Test
        public void proneEmptyColumnDifferentRowSizes() {
            String[][] input = {
                    {"aa", "bb"},
                    {"cc", "dd", "", "ff"},
                    {"gg", "hh", "", "jj"},
                    {"kk"},
            };
            String[][] expected = {
                    {"aa", "bb", ""},
                    {"cc", "dd", "ff"},
                    {"gg", "hh", "jj"},
                    {"kk", "", ""},
            };
            runConsumerTest(createPruneBlankColumnConsumer(), input, expected);
        }

        @Test
        public void proneEmptyColumnEmptyInput() {
            String[][] input = {};
            String[][] expected = {};
            runConsumerTest(createPruneBlankColumnConsumer(), input, expected);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class PruneEmptyRowsAndColumnConsumerTests {
        @Test
        public void removeBlankRowsAndColumnsWhenConfigured() {
            String[][] input = {
                    {"", "aa", "", "bb", ""},
                    {"", "", "", "", "", "", ""},
                    {"", "cc", "", "dd"}
            };
            String[][] expected = {
                    {"aa", "bb"},
                    {"cc", "dd"}
            };
            runConsumerTest(createPruneBlankRowsAndColumnConsumer(), input, expected);
        }

    }

    /**
     *
     * @param consumer
     * @param input
     * @param expected
     */
    private void runConsumerTest(StringRowConsumer consumer, String[][] input, String[][] expected) {
        for (String[] inputRow : input) {
            consumer.accept(inputRow != null ? Arrays.asList(inputRow) : null);
        }
        String[][] resultMatrix = consumer.generateMatrix();
        assertArrayEquals(expected, resultMatrix, "Mismatch expected 2D String array comparison");
    }

    private StringRowConsumer createBasicConsumer() {
        return StringRowConsumer.of(StringRowConsumer.BlankRemoval.NONE);
    }
    private StringRowConsumer createPruneBlankRowsConsumer() {
        return StringRowConsumer.of(StringRowConsumer.BlankRemoval.ROWS);
    }
    private StringRowConsumer createPruneBlankColumnConsumer() {
        return StringRowConsumer.of(StringRowConsumer.BlankRemoval.COLUMNS);
    }
    private StringRowConsumer createPruneBlankRowsAndColumnConsumer() {
        return StringRowConsumer.of(StringRowConsumer.BlankRemoval.ROWS_AND_COLUMNS);
    }
}