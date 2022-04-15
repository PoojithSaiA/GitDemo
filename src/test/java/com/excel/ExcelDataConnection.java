package com.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataConnection {

	public ArrayList<String> getData(String testdata, String TestCases, String Deleteprofile) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream fis=new FileInputStream("C:\\Users\\User\\Documents\\demodata.xlsx");
		ArrayList<String> al =new ArrayList<String>();
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		int sheets=workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase(testdata)) {
				XSSFSheet sheet= workbook.getSheetAt(i);
				Iterator<Row> rows= sheet.rowIterator();
				Row firstrow= rows.next();
				Iterator<Cell> ce=firstrow.cellIterator();
				int k=0;
				int column=0;
				while(ce.hasNext()) {
					Cell c=ce.next();
					if(c.getStringCellValue().equalsIgnoreCase(TestCases)) {
						column=k;
					}k++;
				}
				while(rows.hasNext()) {
					Row r=rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(Deleteprofile)) {
						Iterator<Cell> cv= r.cellIterator();
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
