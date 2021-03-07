package com.example.readpoi;

import android.text.TextUtils;
import android.util.Log;

import org.apache.poi.hpsf.ReadingNotSupportedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import java.util.HashMap;
import java.util.Map;


public class ErmaoExcelProcessor extends ExcelProcessor {

    Map<String, Double> mMoneyMap;

    public ErmaoExcelProcessor(String filePath) {
        super(filePath);
        mMoneyMap = new HashMap<>();
    }

    @Override
    protected void handleSheet(Sheet sheet) {
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum >= 1) {
            for (int i = firstRowNum; i < lastRowNum; i++) {
                Row row = sheet.getRow(i);
                Cell nameCell = row.getCell(2);
                Cell cellbreakFirst = row.getCell(4);
                Cell cellDinner = row.getCell(6);
                try {
                    String name = nameCell.getStringCellValue();
                    if (!TextUtils.isEmpty(name)) {
                        Double moneyBreakFirst = 0D;
                        Double moneyDinner = 0D;
                        if (cellbreakFirst != null) {
                            try {
                                moneyBreakFirst = cellbreakFirst.getNumericCellValue();
                            } catch (Exception e) {
                                if (!TextUtils.isEmpty(cellbreakFirst.getStringCellValue())) {
                                    moneyBreakFirst = Double.valueOf(cellbreakFirst.getStringCellValue());
                                }
                            }

                        }
                        if (cellDinner != null) {
                            try {
                                moneyDinner = cellDinner.getNumericCellValue();
                            } catch (Exception e) {
                                if (!TextUtils.isEmpty(cellDinner.getStringCellValue())) {
                                    moneyDinner = Double.valueOf(cellDinner.getStringCellValue());
                                }
                            }

                        }
                        Double dayMoney = moneyBreakFirst + moneyDinner;
                        if (mMoneyMap.get(name) == null) {
                            mMoneyMap.put(name, dayMoney);
                        } else {
                            Double money = mMoneyMap.get(name) + dayMoney;
                            mMoneyMap.put(name, money);
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public Map<String,Double> getMoneyMap(){
        return mMoneyMap;
    }
}
