/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.io;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Set;
import java.util.zip.GZIPInputStream;

public class InputStreamGenerator {
    private static final Set<String> VALID_URL_SCHEMES =
            Set.of("http", "https", "ftp", "file");
    private static final int CONNECTION_TIMEOUT = 20000;
    // some websites require a userAgent value set.
    private static final String USER_AGENT_VALUE = "javaClient/" + System.getProperty("java.version");

    public InputStream getInputStream(File inputFile) throws IOException {
        return getInputStream(inputFile != null ? inputFile.toPath() : null);
    }

    public InputStream getInputStream(Path inputFile) throws IOException {
        if (inputFile == null) {
            throw new IllegalArgumentException("Must provide an input file.");
        }
        else if (!Files.exists(inputFile)) {
            throw new FileNotFoundException(String.format("Invalid Excel file path: %s", inputFile.toAbsolutePath()));
        }
        else if (Files.isDirectory(inputFile)) {
            throw new IllegalArgumentException("The input file is a directory.");
        }
        return new BufferedInputStream( Files.newInputStream(inputFile) );
    }

    public InputStream getInputStream(URL url) throws IOException {
        if (url == null) {
            throw new IllegalArgumentException("Must provide an input url.");
        }
        String urlProtocol = url.getProtocol();
        if (! VALID_URL_SCHEMES.contains(urlProtocol)) {
            throw new IllegalArgumentException(String.format("URL has an unsupported protocol: %s", urlProtocol));
        }

        if (urlProtocol.equalsIgnoreCase("file")) {
            return getInputStream( Paths.get(url.getPath()) );
        }

        // Could switch to an httpClient in the future.
        URLConnection connection = openConnection(url);
        connection.setConnectTimeout(CONNECTION_TIMEOUT);
        connection.setReadTimeout(CONNECTION_TIMEOUT);
        connection.setRequestProperty("User-Agent", USER_AGENT_VALUE);
        connection.setRequestProperty("Accept-Encoding", "gzip");
        String encoding = connection.getContentEncoding();
        InputStream urlInputStream = connection.getInputStream();
        if ("gzip".equalsIgnoreCase(encoding)) {
            urlInputStream = new GZIPInputStream(urlInputStream);
        }
        return new BufferedInputStream(urlInputStream);
    }

    // scoped to allow for mock testing
    protected URLConnection openConnection(URL url) throws IOException {
        return url.openConnection();
    }
}
