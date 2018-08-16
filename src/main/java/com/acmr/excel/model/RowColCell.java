package com.acmr.excel.model;

import java.io.Serializable;

public class RowColCell implements Serializable{
	// 行
	private Integer row;
	// 列
	private Integer col;
	// 表格对应的ID
	private String cellId;

	public Integer getRow() {
		return row;
	}

	public void setRow(Integer row) {
		this.row = row;
	}

	public Integer getCol() {
		return col;
	}

	public void setCol(Integer col) {
		this.col = col;
	}

	public String getCellId() {
		return cellId;
	}

	public void setCellId(String cellId) {
		this.cellId = cellId;
	}

}
