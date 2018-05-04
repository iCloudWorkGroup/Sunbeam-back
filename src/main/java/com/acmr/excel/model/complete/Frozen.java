package com.acmr.excel.model.complete;

import java.io.Serializable;

public class Frozen implements Serializable{
	private Integer row;
	private Integer col;
	private Integer viewRow;
	private Integer viewCol;
	private String colAlias;
	private String rowAlias;
	
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

	public Integer getViewRow() {
		return viewRow;
	}

	public void setViewRow(Integer viewRow) {
		this.viewRow = viewRow;
	}

	public Integer getViewCol() {
		return viewCol;
	}

	public void setViewCol(Integer viewCol) {
		this.viewCol = viewCol;
	}

	public String getColAlias() {
		return colAlias;
	}

	public void setColAlias(String colAlias) {
		this.colAlias = colAlias;
	}

	public String getRowAlias() {
		return rowAlias;
	}

	public void setRowAlias(String rowAlias) {
		this.rowAlias = rowAlias;
	}
	
}
