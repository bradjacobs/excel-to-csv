/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.io;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.io.TempDir;
import org.mockito.Mockito;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.zip.GZIPOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class InputStreamGeneratorTest {

    private static final InputStreamGenerator inputStreamGenerator = new InputStreamGenerator();

    @TempDir
    private Path tempDir;

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class HappyPathTests {
        @Test
        public void pathToInputStream() throws IOException {
            String testText = "hello world";
            Path tempFile = createTempTextFile(testText);
            try (InputStream is = inputStreamGenerator.getInputStream(tempFile)) {
                assertEquals(testText, readStream(is));
            }
        }

        @Test
        public void fileToInputStream() throws IOException {
            String testText = "hello world";
            Path tempFile = createTempTextFile(testText);
            try (InputStream is = inputStreamGenerator.getInputStream(tempFile.toFile())) {
                assertEquals(testText, readStream(is));
            }
        }

        // give a file as a URL, and expect valid input stream.
        @Test
        public void fileAsUrlInputStream() throws IOException {
            String testText = "hello world";
            Path tempFile = createTempTextFile(testText);
            URL url = tempFile.toUri().toURL();
            try (InputStream is = inputStreamGenerator.getInputStream(url)) {
                assertEquals(testText, readStream(is));
            }
        }

        // test getting inputStream from URL via mocks.
        @Test
        public void urlToInputStream() throws IOException {
            String testText = "hello world";
            InputStreamGenerator streamGenerator =
                    createUrlMockedInputStreamGenerator(testText, false);

            InputStream resultInputStream = streamGenerator
                    .getInputStream(new URL("https://myexample.com/file.txt"));
            assertEquals(testText, readStream(resultInputStream));
        }

        // test getting inputStream from URL with gzip via mocks.
        @Test
        public void urlGzipToInputStream() throws IOException {
            String testText = "hello world";
            InputStreamGenerator streamGenerator =
                    createUrlMockedInputStreamGenerator(testText, true);

            InputStream resultInputStream = streamGenerator
                    .getInputStream(new URL("https://myexample.com/file.txt"));
            assertEquals(testText, readStream(resultInputStream));
        }

        @Test
        public void getUrlConnection() throws IOException {
            // silly test to enforce code coverage stats,
            // as this method used by the Mocks doesn't always register.
            URLConnection connection = inputStreamGenerator.openConnection(new URL("http://fakefake"));
            assertNotNull(connection);
        }

        /**
         * Creates an InputStreamGenerator where given a URL
         * uses mocks to return a given inputStream for the given content
         * @param content content string to use for inputStream
         * @param useGzip true to mock with gzip
         * @return InputStreamGenerator
         * @throws IOException exception
         */
        private InputStreamGenerator createUrlMockedInputStreamGenerator(String content, boolean useGzip) throws IOException {
            InputStreamGenerator streamGenerator = Mockito.spy(new InputStreamGenerator());
            URLConnection mockConnection = mock(URLConnection.class);
            Mockito.doReturn(mockConnection)
                    .when(streamGenerator)
                    .openConnection(Mockito.any());

            byte[] contentBytes = content.getBytes(StandardCharsets.UTF_8);

            if (useGzip) {
                when(mockConnection.getContentEncoding()).thenReturn("gzip");

                ByteArrayOutputStream byteStream = new ByteArrayOutputStream();
                try (GZIPOutputStream gzip = new GZIPOutputStream(byteStream)) {
                    gzip.write(contentBytes);
                }
                contentBytes = byteStream.toByteArray();
            }

            when(mockConnection.getInputStream())
                    .thenReturn(new ByteArrayInputStream(contentBytes));
            return streamGenerator;
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class ErrorHandlingTests {

        @Test
        public void nullPathInput() {
            Exception exception = assertThrows(IllegalArgumentException.class,
                    () -> inputStreamGenerator.getInputStream((Path)null));
            assertEquals("Must provide an input file.",
                    exception.getMessage());
        }

        @Test
        public void nullFileInput() {
            Exception exception = assertThrows(IllegalArgumentException.class,
                    () -> inputStreamGenerator.getInputStream((File)null));
            assertEquals("Must provide an input file.",
                    exception.getMessage());
        }

        @Test
        public void fileNotExistPathInput() {
            Path fakeFile = Path.of("fake/file.xlsx").toAbsolutePath();
            Exception exception = assertThrows(FileNotFoundException.class,
                    () -> inputStreamGenerator.getInputStream(fakeFile));
            assertEquals("Invalid Excel file path: " + fakeFile,
                    exception.getMessage());
        }

        @Test
        public void isDirectoryPathInput() {
            Path dir = Path.of(".").toAbsolutePath();
            Exception exception = assertThrows(IllegalArgumentException.class,
                    () -> inputStreamGenerator.getInputStream(dir));
            assertEquals("The input file cannot be a directory.",
                    exception.getMessage());
        }

        @Test
        public void nullUrlInput() {
            Exception exception = assertThrows(IllegalArgumentException.class,
                    () -> inputStreamGenerator.getInputStream((URL)null));
            assertEquals("Must provide an input url.",
                    exception.getMessage());
        }

        @Test
        public void invalidSchemeUrlInput() throws MalformedURLException {
            URL url = new URL("jar:file:/path/abc.jar!/foo/file.xlsx");
            Exception exception = assertThrows(IllegalArgumentException.class,
                    () -> inputStreamGenerator.getInputStream(url));
            assertEquals("URL has an unsupported protocol: jar",
                    exception.getMessage());
        }
    }

    private String readStream(InputStream input) throws IOException {
        return new String(input.readAllBytes(), StandardCharsets.UTF_8);
    }

    private static final String TEMP_TEXT_FILE_PREFIX = "temp_";
    private static final String TEXT_FILE_EXTENSION = ".txt";

    private Path createTempTextFile(String content) {
        try {
            Path tempTextFile = Files.createTempFile(
                    tempDir,
                    TEMP_TEXT_FILE_PREFIX,
                    TEXT_FILE_EXTENSION
            );
            Files.writeString(tempTextFile, content);
            return tempTextFile;
        }
        catch (IOException e) {
            throw new UncheckedIOException("Unable to create temp text file.", e);
        }
    }
}
