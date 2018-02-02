package com.acmr.excel.model.mongo;

import java.io.Serializable;

import acmr.excel.pojo.ExcelRow;

public class MExcelRow implements Serializable {
	private int rowSort;
	private ExcelRow excelRow;

	public int getRowSort() {
		return rowSort;
	}

	public void setRowSort(int rowSort) {
		this.rowSort = rowSort;
	}

	public ExcelRow getExcelRow() {
		return excelRow;
	}

	public void setExcelRow(ExcelRow excelRow) {
		this.excelRow = excelRow;
	}

}
