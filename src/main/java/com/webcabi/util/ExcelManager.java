package com.webcabi.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelManager {
	private Workbook workbook;
	
	public void open(File file) {
		try {
			FileInputStream input = new FileInputStream(file);
			this.workbook = new XSSFWorkbook(input);
		} catch(FileNotFoundException ex) {
			ex.printStackTrace();
		} catch(IOException ex) {
			ex.printStackTrace();
		}
	}
	
	public void close() {
		try {
			this.workbook.close();
		} catch(IOException ex) {
			ex.printStackTrace();
		}
	}
}
