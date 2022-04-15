package com.excel.practice;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelConnection {

	public ArrayList<String> getData(String string, String string2, String string3) throws IOException {
		FileInputStream fis=new FileInputStream("C:\\Users\\User\\Documents\\demodata.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		int sc=workbook.getNumberOfSheets();
		ArrayList<String> al=new ArrayList<String>();
		for (int i = 0; i <sc; i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase(string)) {
				XSSFSheet sheet =workbook.getSheetAt(i);
				Iterator<Row>rows= sheet.rowIterator();
				Row firstrow=rows.next();
				Iterator<Cell> ce=firstrow.cellIterator();
				int k=0;
				int column=0;
				while(ce.hasNext()) {
					Cell c=ce.next();
					if(c.getStringCellValue().equalsIgnoreCase(string2)) {
					column=k;
					}
					k++;
				}
				while(rows.hasNext()) {
					Row r=rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(string3)) {
						Iterator<Cell> cv=r.cellIterator();
						while(cv.hasNext()) {
							Cell c=cv.next();
							if(c.getCellType()==CellType.STRING) {
								al.add(c.getStringCellValue());
							}
							else {
								al.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}
				}
			}
		}
		return al;
	}

}
