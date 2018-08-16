package com.acmr.excel.model.mongo;

import java.io.Serializable;

public class MRowColCell implements Serializable{
	/* 行id */
	private String row;
	/* 列id */
	private String col;
	/* 单元格id */
	private String cellId;
	private String sheetId;

	public String getRow() {
		return row;
	}

	public void setRow(String row) {
		this.row = row;
	}

	public String getCol() {
		return col;
	}

	public void setCol(String col) {
		this.col = col;
	}

	public String getCellId() {
		return cellId;
	}

	public void setCellId(String cellId) {
		this.cellId = cellId;
	}

	public String getSheetId() {
		return sheetId;
	}

	public void setSheetId(String sheetId) {
		this.sheetId = sheetId;
	}

}
