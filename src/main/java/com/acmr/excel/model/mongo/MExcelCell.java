package com.acmr.excel.model.mongo;

import java.io.Serializable;

import acmr.excel.pojo.ExcelCell;

public class MExcelCell implements Serializable {

	private String id;
	private String rowId;// 开始行
	private String colId;// 开始列
	private int rowspan;// 合并几行,判断是否为合并单元格,计算完整的合并信息,当变化的时候，直接+、-就可以，不用遍历所有包含对象
	private int colspan;// 合并几列

	private ExcelCell excelCell;

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

	public int getRowspan() {
		return rowspan;
	}

	public void setRowspan(int rowspan) {
		this.rowspan = rowspan;
	}

	public int getColspan() {
		return colspan;
	}

	public void setColspan(int colspan) {
		this.colspan = colspan;
	}

	public ExcelCell getExcelCell() {
		return excelCell;
	}

	public void setExcelCell(ExcelCell excelCell) {
		this.excelCell = excelCell;
	}

}
