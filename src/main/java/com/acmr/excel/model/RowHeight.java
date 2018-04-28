package com.acmr.excel.model;

import java.io.Serializable;

public class RowHeight implements Serializable {
	private int row;
	private int offset;

	public int getOffset() {
		return offset;
	}

	public void setOffset(int offset) {
		this.offset = offset;
	}

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}
	
}
