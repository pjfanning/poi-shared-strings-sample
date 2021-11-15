package com.github.pjfanning.poi.sample;

import com.github.pjfanning.poi.xssf.streaming.SXSSFFactory;
import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.poi.ss.usermodel.*;
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
        SXSSFFactory sxssfFactory = new SXSSFFactory()
                .enableTempFileComments(false) //true doesn't work properly (under investigation)
                .encryptTempFiles(true);
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(sxssfFactory),
                SXSSFWorkbook.DEFAULT_WINDOW_SIZE, true, true)) {
            wb.setZip64Mode(Zip64Mode.Always);
            SXSSFSheet sheet = wb.createSheet("SheetA");
            for (int r = 0; r < ROWS; r++) {
                SXSSFRow row = sheet.createRow(r);
                for (int c = 0; c < COLUMNS; c++) {
                    SXSSFCell cell = row.createCell(c);
                    cell.setCellValue(UUID.randomUUID().toString());
                }
                if (r == 0) {
                    addComment(row.getCell(0), "poi-user", "added by StreamingWrite.java");
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

    private static void addComment(Cell cell, String author, String commentText) {
        Sheet sheet = cell.getRow().getSheet();
        CreationHelper factory = sheet.getWorkbook().getCreationHelper();

        ClientAnchor anchor = factory.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex() + 1); //the box of the comment starts at this given column...
        anchor.setCol2(cell.getColumnIndex() + 3); //...and ends at that given column
        anchor.setRow1(cell.getRowIndex() + 1); //one row below the cell...
        anchor.setRow2(cell.getRowIndex() + 5); //...and 4 rows high

        Drawing drawing = sheet.createDrawingPatriarch();
        Comment comment = drawing.createCellComment(anchor);
        //set the comment text and author
        comment.setString(factory.createRichTextString(commentText));
        comment.setAuthor(author);

        cell.setCellComment(comment);
    }
}
