package com.demo.ExcelProject;

import java.io.IOException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
		File f = new File("Demo.xlsx");
        FileInputStream fis = new FileInputStream(f);
        XSSFWorkbook excelWorkbook = new XSSFWorkbook(fis);
        XSSFSheet excelSheet = excelWorkbook.getSheetAt(0);
        XSSFCell cell=excelSheet.getRow(0).createCell(2);//row index - 0,column index-2
        cell.setCellType(CellType.STRING);
        cell.setCellValue("Pass");
        FileOutputStream fos = new FileOutputStream("Demo.xlsx");
        excelWorkbook.write(fos);
        fos.close();
        fis.close();
		}catch(Exception e) {
			e.printStackTrace();
		}

	}

}
