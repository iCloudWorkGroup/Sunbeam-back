package com.acmr.excel.test;

import java.lang.reflect.Field;

import com.acmr.excel.model.mongo.MRow;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;

public class TestUtil {
	public static ExcelBook createNewExcel() {
		ExcelBook excelBook = new ExcelBook();
		ExcelSheet sheet = new ExcelSheet();
		for (int i = 1; i < 27; i++) {
			ExcelColumn column = sheet.addColumn();
			column.setWidth(69);
		}
		for (int i = 1; i < 101; i++) {
			ExcelRow row = sheet.addRow();
			row.setHeight(19);
		}
		excelBook.getSheets().add(sheet);
		return excelBook;
	}
	public static void main(String[] args) {
		MRow mr = new MRow();
		mr.getProps().getContent().setColor("123");
		try {
			Field fRow = mr.getProps().getContent().getClass()
					.getDeclaredField("color");
			fRow.setAccessible(true);
			Object value = fRow.get(mr.getProps().getContent());
			System.out.println(value);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
}
