package com.excel;

import java.io.IOException;
import java.util.ArrayList;

public class ExcelMain {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		ExcelDataConnection xl=new ExcelDataConnection();
		ArrayList<String> al=xl.getData("testdata","TestCases","Delete profile");
		System.out.println(al);
	}

}
