package main;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class ExcelBreaker {
    File zipExcelFIle;
    ArrayList<File> unzipList = null;
    String rootDirectory = null;
    String brakedFile = null;

    public ExcelBreaker(File excelFile) {
        this.zipExcelFIle = excelFile;

        try {
            unzip();
            List<String> fileNames = new ArrayList<>(Arrays.asList(rootDirectory.split("\\.")));
            fileNames.remove(fileNames.size()-1);
            brakedFile = String.join(".", fileNames);

            compress(rootDirectory, brakedFile);
        } catch (Throwable var3) {
            var3.printStackTrace();
        }

    }

    public static void compress(String path, String outputFileName) throws Throwable {
        System.out.println(path);
        System.out.println(outputFileName);
        File file = new File(path);

        if (!file.exists()) throw new Exception("압축할 폴더를 선택해주세요");
        FileOutputStream fos = null;
        ZipOutputStream zos = null;

        try {
            fos = new FileOutputStream(new File(outputFileName));
            zos = new ZipOutputStream(fos);

            searchDirectory(file, zos);
        } catch (Throwable var10) {
            throw var10;
        } finally {
            if (zos != null) {
                zos.close();
            }

            if (fos != null) {
                fos.close();
            }

        }
    }

    private static void searchDirectory(File file, ZipOutputStream zos) throws Throwable {
        searchDirectory(file, file.getPath(), zos);
    }

    private static void searchDirectory(File file, String root, ZipOutputStream zos) throws Exception {
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            File[] var7 = files;
            int var6 = files.length;

            for(int var5 = 0; var5 < var6; ++var5) {
                File f = var7[var5];
                searchDirectory(f, root, zos);
            }
        } else {
            compressZip(file, root, zos);
        }

    }

    private static void compressZip(File file, String root, ZipOutputStream zos) throws Exception {
        FileInputStream fis = null;

        try {
            String zipName = file.getPath().replace(root + "\\", "");
            fis = new FileInputStream(file);
            ZipEntry zipentry = new ZipEntry(zipName);
            zos.putNextEntry(zipentry);
            int length = (int)file.length();
            byte[] buffer = new byte[length];
            fis.read(buffer, 0, length);
            zos.write(buffer, 0, length);
            zos.closeEntry();
        } catch (Throwable var11) {
            throw var11;
        } finally {
            if (fis != null) {
                fis.close();
            }

        }

    }

    private void unzip() throws Throwable {
        FileInputStream fis = null;
        ZipInputStream zis = null;
        ZipEntry zipentry = null;
        this.unzipList = new ArrayList();

        try {
            fis = new FileInputStream(this.zipExcelFIle);
            zis = new ZipInputStream(fis);

            while((zipentry = zis.getNextEntry()) != null) {
                String filename = zipentry.getName();
                this.rootDirectory = this.zipExcelFIle.getAbsolutePath() + "tmp";
                File file = new File(this.rootDirectory, filename);
                this.unzipList.add(file);
                if (zipentry.isDirectory()) {
                    file.mkdirs();
                } else {
                    this.createFile(file, zis);
                    if (filename.contains("vbaProject.bin")) {
                        replaceText(file, "DPB", "DPx");
                    }

                    if (filename.contains("workbook.xml") && !filename.contains("rels")) {
                        this.removeNode(file, "workbookProtection");
                    }

                    if (filename.contains("worksheets/sheet")) {
                        this.removeNode(file, "sheetProtection");
                    }
                }
            }
        } catch (Throwable var9) {
            throw var9;
        } finally {
            if (zis != null) {
                zis.close();
            }

            if (fis != null) {
                fis.close();
            }

        }

    }

    private void createFile(File file, ZipInputStream zis) throws Throwable {
        File parentDir = new File(file.getParent());
        if (!parentDir.exists()) {
            parentDir.mkdirs();
        }

        FileOutputStream fos = null;

        try {
            fos = new FileOutputStream(file);
            byte[] buffer = new byte[256];
            boolean var6 = false;

            int size;
            while((size = zis.read(buffer)) > 0) {
                fos.write(buffer, 0, size);
            }

        } catch (Throwable var7) {
            throw var7;
        }
    }

    private void replaceText(File target, String text, String replacement) throws IOException {
        String origin_text = readTextFile(target).replaceAll(text, replacement);
        BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(target),  StandardCharsets.UTF_8));

        writer.write(origin_text);
        writer.flush();
        writer.close();
    }

    private void removeNode(File target, String nodeName) throws IOException {
        String origin_text = readTextFile(target).replaceAll("<" + nodeName + "[\\w\\d =\"\\-\\+\\/]*>", "");
        BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(target), StandardCharsets.UTF_8));
        writer.write(origin_text);
        writer.flush();
        writer.close();
    }

    private String readTextFile(File target) throws IOException {
        String origin_text = "";
        BufferedReader reader = null;

        try {
            reader = new BufferedReader(new InputStreamReader(new FileInputStream(target),  StandardCharsets.UTF_8));
        } catch (FileNotFoundException var7) {
            var7.printStackTrace();
        }

        for(String schar = ""; (schar = reader.readLine()) != null; origin_text = origin_text + schar) {}

        return origin_text;
    }
}