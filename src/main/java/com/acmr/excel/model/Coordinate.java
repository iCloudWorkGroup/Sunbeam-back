package com.acmr.excel.model;

import java.io.Serializable;

public class Coordinate implements Serializable {

	private int startCol;
	private int startRow;
	private int endCol;
	private int endRow;
	public int getStartCol() {
		return startCol;
	}
	public void setStartCol(int startCol) {
		this.startCol = startCol;
	}
	public int getStartRow() {
		return startRow;
	}
	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}
	public int getEndCol() {
		return endCol;
	}
	public void setEndCol(int endCol) {
		this.endCol = endCol;
	}
	public int getEndRow() {
		return endRow;
	}
	public void setEndRow(int endRow) {
		this.endRow = endRow;
	}

	
}
