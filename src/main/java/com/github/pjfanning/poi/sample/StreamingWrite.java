package com.github.pjfanning.poi.sample;

import com.github.pjfanning.poi.xssf.streaming.SXSSFFactory;
import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.UUID;

public class StreamingWrite {
    private static String OUTPUT_FILENAME = "StreamingWrite.xlsx";
    private static int ROWS = 1000;
    private static int COLUMNS = 10;

    public static void main(String[] args) {
        // the SXSSFFactory ensures that TempFileSharedStringsTable is used instead of the in-memory default
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(new SXSSFFactory().encryptTempFiles(true)),
                SXSSFWorkbook.DEFAULT_WINDOW_SIZE, true, true)) {
            wb.setZip64Mode(Zip64Mode.Always);
            SXSSFSheet sheet = wb.createSheet("SheetA");
            for (int r = 0; r < ROWS; r++) {
                SXSSFRow row = sheet.createRow(r);
                for (int c = 0; c < COLUMNS; c++) {
                    SXSSFCell cell = row.createCell(c);
                    cell.setCellValue(UUID.randomUUID().toString());
                }
            }
            try (FileOutputStream fos = new FileOutputStream(OUTPUT_FILENAME)) {
                System.out.println("Writing xlsx file to " + OUTPUT_FILENAME);
                wb.write(fos);
            }
        } catch (Throwable t) {
            t.printStackTrace();
        }
    }
}
