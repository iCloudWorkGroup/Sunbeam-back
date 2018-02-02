package com.acmr.excel.model.mongo;

import java.io.Serializable;

import acmr.excel.pojo.ExcelColumn;

public class MExcelColumn implements Serializable {
	private int colSort;
	private ExcelColumn excelColumn;

	public int getColSort() {
		return colSort;
	}

	public void setColSort(int colSort) {
		this.colSort = colSort;
	}

	public ExcelColumn getExcelColumn() {
		return excelColumn;
	}

	public void setExcelColumn(ExcelColumn excelColumn) {
		this.excelColumn = excelColumn;
	}

}
