package com.example.readpoi;

import android.os.Build;
import android.webkit.WebSettings;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelProcessor {

    private String mFilePath;
    private HSSFWorkbook mWorkBook;
    private XSSFWorkbook xssfWorkbook;
    public ExcelProcessor(String filePath){
        mFilePath = filePath;
        try {
            if (Build.VERSION.SDK_INT > Build.VERSION_CODES.LOLLIPOP) {
                xssfWorkbook = new XSSFWorkbook(new FileInputStream(mFilePath));
            }else {
                mWorkBook = new HSSFWorkbook(new FileInputStream(mFilePath));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void readSheet() {
        if (Build.VERSION.SDK_INT > Build.VERSION_CODES.LOLLIPOP) {
            for (int sheetIndex = 0; sheetIndex < xssfWorkbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = xssfWorkbook.getSheetAt(sheetIndex);
                handleSheet(sheet);
            }
        }else {
            for (int sheetIndex = 0; sheetIndex < mWorkBook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = mWorkBook.getSheetAt(sheetIndex);
                handleSheet(sheet);
            }
        }

    }

    protected void handleSheet(Sheet sheet){

    }



}
