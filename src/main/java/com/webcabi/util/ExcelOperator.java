package com.webcabi.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperator {
    private Workbook workbook;

    public final void open(final File file) {
        try {
            FileInputStream input = new FileInputStream(file);
            this.workbook = new XSSFWorkbook(input);
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
    
    public final void open(final String file) {
        this.open(new File(file));
    }

    public final void close() {
        try {
            this.workbook.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
    
    public final ArrayList<String> readColumnWithString(final String sheetName,
            final int colName, final int offset) {
        return this.readColumnWithString(sheetName, colName, offset, 0, false);
    }
    
    public final ArrayList<String> readColumnWithString(final String sheetName,
            final int colNum, final int offset, int size, final boolean force) {
        ArrayList<String> res = new ArrayList<String>();
        int maxRowNum = offset + size;
        
        Sheet sheet = workbook.getSheet(sheetName);
        int lastRowNum = sheet.getLastRowNum() + 1;
        if (size == 0) {
            maxRowNum = lastRowNum;
        }
        if (!force && maxRowNum >= lastRowNum) {
            maxRowNum = lastRowNum;
        }
        
        for (int r = offset; r < maxRowNum; r++) {
            Row row = sheet.getRow(r);
            String value = getStringByCell(row.getCell(colNum));
            
            if (null == value) {
                if (force) {
                    res.add("");
                } else {
                    break;
                }
            } else {
                res.add(value);
            }
        }
        
        return res;
    }
    
    /*
     * Get value of cell with String.
     */
    private String getStringByCell(final Cell cell) {
        if (null != cell) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    return Boolean.toString(cell.getBooleanCellValue()).trim();
                case Cell.CELL_TYPE_FORMULA:
                    Workbook parent = cell.getSheet().getWorkbook();
                    CreationHelper helper = parent.getCreationHelper();
                    FormulaEvaluator evaluator = helper.createFormulaEvaluator();
                    return getStringByCell(evaluator.evaluateInCell(cell)).trim();
                case Cell.CELL_TYPE_NUMERIC:
                    int i = (int) cell.getNumericCellValue();
                    return Integer.toString(i);
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue().trim();
                default:
                    return "";
            }
        }
        
        return null;
    }
}
