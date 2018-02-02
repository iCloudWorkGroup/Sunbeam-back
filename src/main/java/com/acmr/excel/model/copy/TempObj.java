package com.acmr.excel.model.copy;

import acmr.excel.pojo.ExcelCell;

public class TempObj {
	private int row;
	private int col;
	private ExcelCell excelCell;

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public int getCol() {
		return col;
	}

	public void setCol(int col) {
		this.col = col;
	}

	public ExcelCell getExcelCell() {
		return excelCell;
	}

	public void setExcelCell(ExcelCell excelCell) {
		this.excelCell = excelCell;
	}

}
