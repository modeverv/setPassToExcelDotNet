package com.modeverv.poi;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public final class PoiDecryptChecker {
    private PoiDecryptChecker() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 4 || !"decrypt".equalsIgnoreCase(args[0])) {
            printUsageAndExit();
            return;
        }

        var encryptedPath = args[1];
        var outputPath = args[2];
        var password = args[3];

        decrypt(encryptedPath, outputPath, password);
    }

    private static void decrypt(String encryptedPath, String outputPath, String password) throws Exception {
        try (var input = new FileInputStream(encryptedPath);
             var fs = new POIFSFileSystem(input)) {
            var info = new EncryptionInfo(fs);
            var decryptor = Decryptor.getInstance(info);
            if (!decryptor.verifyPassword(password)) {
                throw new IllegalArgumentException("Invalid password");
            }

            try (InputStream dataStream = decryptor.getDataStream(fs);
                 OutputStream out = new FileOutputStream(outputPath)) {
                copy(dataStream, out);
            }
        }
    }

    private static void copy(InputStream input, OutputStream output) throws IOException {
        var buffer = new byte[8192];
        int read;
        while ((read = input.read(buffer)) != -1) {
            output.write(buffer, 0, read);
        }
    }

    private static void printUsageAndExit() {
        System.err.println("Usage: java -jar poi-decrypt-checker.jar decrypt <encrypted.xlsx> <output.xlsx> <password>");
        System.exit(2);
    }
}

