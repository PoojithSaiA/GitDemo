package com.excel.practice;

import java.io.IOException;
import java.util.ArrayList;

public class ExcelMain {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		ExcelConnection ec=new ExcelConnection();
		ArrayList<String> al=ec.getData("testdata","TestCases","Purchase");
		System.out.println(al);
	}

}
