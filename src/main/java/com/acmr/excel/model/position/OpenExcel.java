package com.acmr.excel.model.position;

public class OpenExcel {
	private String excelId;
	private int top;
	private int bottom;
	private int colBegin;
	private int colEnd;

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	
	public int getTop() {
		return top;
	}

	public void setTop(int top) {
		this.top = top;
	}

	public int getBottom() {
		return bottom;
	}

	public void setBottom(int bottom) {
		this.bottom = bottom;
	}

	public int getColBegin() {
		return colBegin;
	}

	public void setColBegin(int colBegin) {
		this.colBegin = colBegin;
	}

	public int getColEnd() {
		return colEnd;
	}

	public void setColEnd(int colEnd) {
		this.colEnd = colEnd;
	}

}
