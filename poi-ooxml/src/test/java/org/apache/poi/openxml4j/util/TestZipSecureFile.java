/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.openxml4j.util;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.hssf.HSSFTestDataSamples;
import org.apache.poi.openxml4j.opc.internal.InvalidZipException;
import org.apache.poi.openxml4j.opc.internal.ZipHelper;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.XSSFTestDataSamples;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Enumeration;

import static org.junit.jupiter.api.Assertions.*;

class TestZipSecureFile {
    @Test
    void testThresholdInputStream() throws Exception {
        // This fails in Java 10 because our reflection injection of the ThresholdInputStream causes a
        // ClassCastException in ZipFile now
        // The relevant change in the JDK is http://hg.openjdk.java.net/jdk/jdk10/rev/85ea7e83af30#l5.66

        try (ZipFile thresholdInputStream = new ZipFile(XSSFTestDataSamples.getSampleFile("template.xlsx"))) {
            try (ZipSecureFile secureFile = new ZipSecureFile(XSSFTestDataSamples.getSampleFile("template.xlsx"))) {
                Enumeration<? extends ZipArchiveEntry> entries = thresholdInputStream.getEntries();
                while (entries.hasMoreElements()) {
                    ZipArchiveEntry entry = entries.nextElement();

                    try (InputStream inputStream = secureFile.getInputStream(entry)) {
                        assertTrue(IOUtils.toByteArray(inputStream).length > 0);
                    }
                }
            }
        }
    }

    @Test
    void testSettingMaxEntrySizeAsNegative() {
        assertThrows(IllegalArgumentException.class, () -> ZipSecureFile.setMaxEntrySize(-1));
    }

    @Test
    void testSettingMaxEntrySizeAs8Gb() {
        long approx8Gb = ZipSecureFile.MAX_ENTRY_SIZE * 2;
        try {
            ZipSecureFile.setMaxEntrySize(approx8Gb);
            assertEquals(approx8Gb, ZipSecureFile.getMaxEntrySize());
        } finally {
            ZipSecureFile.setMaxEntrySize(ZipSecureFile.MAX_ENTRY_SIZE);
        }
    }

    @Test
    void testSettingMaxTextSizeAsNegative() {
        assertThrows(IllegalArgumentException.class, () -> ZipSecureFile.setMaxTextSize(-1));
    }

    @Test
    void testSettingMaxTextSizeAs8GChars() {
        long approx8G = ZipSecureFile.MAX_ENTRY_SIZE * 2;
        try {
            ZipSecureFile.setMaxTextSize(approx8G);
            assertEquals(approx8G, ZipSecureFile.getMaxTextSize());
        } finally {
            ZipSecureFile.setMaxTextSize(ZipSecureFile.DEFAULT_MAX_TEXT_SIZE);
        }
    }

    @Test
    void testDuplicateEntryNamesFile() {
        File file = XSSFTestDataSamples.getSampleFile("duplicate-file.xlsx");
        InvalidZipException ex = assertThrows(InvalidZipException.class, () -> {
            try (ZipSecureFile zf = new ZipSecureFile(file)) {
                // expect exception here
            }
        });
        assertEquals("Input file contains more than 1 entry with the name xl/worksheets/sheet1.xml", ex.getMessage());

        ex = assertThrows(InvalidZipException.class, () -> {
            try (ZipSecureFile zf = new ZipSecureFile(file.getAbsolutePath())) {
                // expect exception here
            }
        });
        assertEquals("Input file contains more than 1 entry with the name xl/worksheets/sheet1.xml", ex.getMessage());
    }

    @Test
    void testDuplicateEntryNamesStream() {
        InvalidZipException ex = assertThrows(InvalidZipException.class, () -> {
            try (InputStream is = HSSFTestDataSamples.openSampleFileStream("duplicate-file.xlsx");
                 ZipInputStreamZipEntrySource zes = new ZipInputStreamZipEntrySource(ZipHelper.openZipStream(is))) {
                // expect exception here
            }
        });
        assertEquals("Input file contains more than 1 entry with the name xl/worksheets/sheet1.xml", ex.getMessage());
    }

    @Test
    void testDuplicateEntryNamesDifferentCase() throws IOException {
        assertRejected("Input file contains more than 1 entry with the name XL/Worksheets/Sheet1.xml",
                "[Content_Types].xml", "xl/worksheets/sheet1.xml", "XL/Worksheets/Sheet1.xml");
    }

    @Test
    void testDuplicateEntryNamesMixedSeparators() throws IOException {
        assertRejected("Input file contains more than 1 entry with the name xl/worksheets\\sheet1.xml",
                "[Content_Types].xml", "xl/worksheets/sheet1.xml", "xl/worksheets\\sheet1.xml");
    }

    @Test
    void testEmptyEntryName() throws IOException {
        assertRejected("Input file contains an entry with an empty name",
                "[Content_Types].xml", "");
    }

    @Test
    void testDistinctEntryNames() throws IOException {
        final byte[] data = createZip("[Content_Types].xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet10.xml");
        final File file = writeTempFile(data);
        try {
            try (ZipSecureFile zf = new ZipSecureFile(file)) {
                assertNotNull(zf.getEntry("xl/worksheets/sheet10.xml"));
            }
        } finally {
            assertTrue(file.delete());
        }
        try (ZipInputStreamZipEntrySource zes = new ZipInputStreamZipEntrySource(
                ZipHelper.openZipStream(toInputStream(data)))) {
            assertNotNull(zes.getEntry("xl/worksheets/sheet1.xml"));
            assertNotNull(zes.getEntry("xl/worksheets/sheet10.xml"));
        }
    }

    private static void assertRejected(String expectedMessage, String... entryNames) throws IOException {
        final byte[] data = createZip(entryNames);
        final File file = writeTempFile(data);
        try {
            InvalidZipException ex = assertThrows(InvalidZipException.class, () -> {
                try (ZipSecureFile zf = new ZipSecureFile(file)) {
                    // expect exception here
                }
            });
            assertEquals(expectedMessage, ex.getMessage());
        } finally {
            assertTrue(file.delete());
        }

        InvalidZipException ex = assertThrows(InvalidZipException.class, () -> {
            try (ZipInputStreamZipEntrySource zes = new ZipInputStreamZipEntrySource(
                    ZipHelper.openZipStream(toInputStream(data)))) {
                // expect exception here
            }
        });
        assertEquals(expectedMessage, ex.getMessage());
    }

    private static byte[] createZip(String... entryNames) throws IOException {
        final UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream();
        try (ZipArchiveOutputStream zos = new ZipArchiveOutputStream(bos)) {
            int i = 0;
            for (String entryName : entryNames) {
                zos.putArchiveEntry(new ZipArchiveEntry(entryName));
                zos.write(("<entry" + (i++) + "/>").getBytes(StandardCharsets.UTF_8));
                zos.closeArchiveEntry();
            }
        }
        return bos.toByteArray();
    }

    private static InputStream toInputStream(byte[] data) {
        return new ByteArrayInputStream(data);
    }

    private static File writeTempFile(byte[] data) throws IOException {
        final File file = TempFile.createTempFile("zip-entry-names", ".zip");
        try (OutputStream os = new FileOutputStream(file)) {
            os.write(data);
        }
        return file;
    }
}
