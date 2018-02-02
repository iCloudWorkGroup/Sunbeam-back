package com.acmr.excel.model.mongo;

import java.io.Serializable;

import acmr.excel.pojo.ExcelCell;

public class MExcelCell implements Serializable {

	private String id;
	private String rowId;
	private String colId;

	private ExcelCell excelCell;

	public ExcelCell getExcelCell() {
		return excelCell;
	}

	public void setExcelCell(ExcelCell excelCell) {
		this.excelCell = excelCell;
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getRowId() {
		return rowId;
	}

	public void setRowId(String rowId) {
		this.rowId = rowId;
	}

	public String getColId() {
		return colId;
	}

	public void setColId(String colId) {
		this.colId = colId;
	}

}
