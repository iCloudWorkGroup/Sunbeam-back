package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class CompleteExcel implements Serializable {
	private String name = "新建excel";
	private List<SheetElement> sheets = new ArrayList<SheetElement>();

	public List<SheetElement> getSheets() {
		return sheets;
	}

	public void setSheets(List<SheetElement> sheets) {
		this.sheets = sheets;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
}
