package com.example.readpoi;

import android.os.Environment;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;

public class ExcelMaker {

    public ExcelMaker(){

    }

    public String write(Map<String, Double> mMoneyMap){
        Workbook wb = new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("汇总");
        int rowIndex = 0;
        for (String name : mMoneyMap.keySet()) {
            int lastNum = rowIndex;
            Row row = sheet.createRow(lastNum);
            row.createCell(0).setCellValue(name);
            row.createCell(1).setCellValue(mMoneyMap.get(name));
            rowIndex++;
        }
        long number = System.currentTimeMillis() / 1000;
        String storageDir = Environment.getExternalStorageDirectory().toString();
        File file = new File(storageDir, "汇总表_" + number + ".xls");
        //将表格输出到文件
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(file);
            wb.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return file.getAbsolutePath();
    }

}
