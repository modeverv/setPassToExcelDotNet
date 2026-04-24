package com.modeverv.poi;

import org.apache.poi.poifs.crypt.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;

public final class PoiDecryptChecker {
    private PoiDecryptChecker() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length < 2) {
            printUsageAndExit();
            return;
        }

        switch (args[0].toLowerCase()) {
            case "decrypt" -> {
                if (args.length != 4) { printUsageAndExit(); return; }
                decrypt(args[1], args[2], args[3]);
            }
            case "encrypt" -> {
                if (args.length != 4) { printUsageAndExit(); return; }
                encrypt(args[1], args[2], args[3]);
            }
            case "create" -> {
                if (args.length != 3) { printUsageAndExit(); return; }
                create(args[1], args[2]);
            }
            default -> printUsageAndExit();
        }
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

    private static void encrypt(String inputPath, String outputPath, String password) throws Exception {
        var encInfo = new EncryptionInfo(EncryptionMode.agile, CipherAlgorithm.aes256, HashAlgorithm.sha512, -1, -1, null);
        var enc = encInfo.getEncryptor();
        enc.confirmPassword(password);

        try (var fs = new POIFSFileSystem()) {
            try (var out = enc.getDataStream(fs)) {
                out.write(Files.readAllBytes(Paths.get(inputPath)));
            }
            try (var fos = new FileOutputStream(outputPath)) {
                fs.writeFilesystem(fos);
            }
        }
    }

    private static void create(String type, String outputPath) throws Exception {
        switch (type.toLowerCase()) {
            case "simple" -> createSimple(outputPath);
            case "formulas" -> createFormulas(outputPath);
            case "styles" -> createStyles(outputPath);
            case "japanese" -> createJapanese(outputPath);
            default -> throw new IllegalArgumentException("Unknown type: " + type + ". Must be simple, formulas, styles, or japanese.");
        }
    }

    private static void createSimple(String outputPath) throws Exception {
        try (var wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Value");
            var row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue("Item A");
            row1.createCell(1).setCellValue(100.0);
            var row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue("Item B");
            row2.createCell(1).setCellValue(200.0);
            try (var out = new FileOutputStream(outputPath)) {
                wb.write(out);
            }
        }
    }

    private static void createFormulas(String outputPath) throws Exception {
        try (var wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("A");
            header.createCell(1).setCellValue("B");
            header.createCell(2).setCellValue("Sum");
            var row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue(10.0);
            row1.createCell(1).setCellValue(20.0);
            row1.createCell(2).setCellFormula("A2+B2");
            var row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue(30.0);
            row2.createCell(1).setCellValue(40.0);
            row2.createCell(2).setCellFormula("A3+B3");
            var row3 = sheet.createRow(3);
            row3.createCell(0).setCellFormula("SUM(A2:A3)");
            row3.createCell(1).setCellFormula("SUM(B2:B3)");
            row3.createCell(2).setCellFormula("SUM(C2:C3)");
            try (var out = new FileOutputStream(outputPath)) {
                wb.write(out);
            }
        }
    }

    private static void createStyles(String outputPath) throws Exception {
        try (var wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");

            var boldFont = wb.createFont();
            boldFont.setBold(true);

            var headerStyle = wb.createCellStyle();
            headerStyle.setFont(boldFont);
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            var header = sheet.createRow(0);
            var h0 = header.createCell(0);
            h0.setCellValue("Category");
            h0.setCellStyle(headerStyle);
            var h1 = header.createCell(1);
            h1.setCellValue("Amount");
            h1.setCellStyle(headerStyle);

            var row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue("Sales");
            row1.createCell(1).setCellValue(1500.0);

            var row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue("Expenses");
            row2.createCell(1).setCellValue(800.0);

            try (var out = new FileOutputStream(outputPath)) {
                wb.write(out);
            }
        }
    }

    private static void createJapanese(String outputPath) throws Exception {
        try (var wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("シート1");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("名前");
            header.createCell(1).setCellValue("値");
            var row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue("項目A");
            row1.createCell(1).setCellValue(100.0);
            var row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue("項目B");
            row2.createCell(1).setCellValue(200.0);
            var row3 = sheet.createRow(3);
            row3.createCell(0).setCellValue("日本語テスト");
            row3.createCell(1).setCellValue(0.0);
            try (var out = new FileOutputStream(outputPath)) {
                wb.write(out);
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
        System.err.println("Usage:");
        System.err.println("  java -jar poi-decrypt-checker.jar decrypt <encrypted.xlsx> <output.xlsx> <password>");
        System.err.println("  java -jar poi-decrypt-checker.jar encrypt <input.xlsx> <output.xlsx> <password>");
        System.err.println("  java -jar poi-decrypt-checker.jar create <type> <output.xlsx>");
        System.err.println("    types: simple, formulas, styles, japanese");
        System.exit(2);
    }
}
