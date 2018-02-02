package com.acmr.excel.convert;

public class ParamVo {
	/**
	 * 起始行
	 */
	private int rowBegin;
	/**
	 * 结束行
	 */
	private int rowEnd;
	/**
	 * 起始列
	 */
	private int colBegin;
	/**
	 * 结束列
	 */
	private int colEnd;
	/**
	 * 类型
	 */
	private String type;

	public int getRowBegin() {
		return rowBegin;
	}

	public void setRowBegin(int rowBegin) {
		this.rowBegin = rowBegin;
	}

	public int getRowEnd() {
		return rowEnd;
	}

	public void setRowEnd(int rowEnd) {
		this.rowEnd = rowEnd;
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

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

}
