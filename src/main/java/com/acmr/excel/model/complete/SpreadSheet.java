package com.acmr.excel.model.complete;

import java.io.Serializable;

public class SpreadSheet implements Serializable {

	private SheetElement sheet = new SheetElement();
	private String name;
	private int sort;
	private String tempHTML;

	public SheetElement getSheet() {
		return sheet;
	}

	public void setSheet(SheetElement sheet) {
		this.sheet = sheet;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getSort() {
		return sort;
	}

	public void setSort(int sort) {
		this.sort = sort;
	}

	public String getTempHTML() {
		return tempHTML;
	}

	public void setTempHTML(String tempHTML) {
		this.tempHTML = tempHTML;
	}
}
