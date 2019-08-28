package com.github.pjfanning.poi.sample;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;

public class FileRead {

    public static void main(String...strings) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook(new File("poc.xlsx"));
        System.out.println(wb.getSheetName(0));
        System.out.println("Done");
    }
}