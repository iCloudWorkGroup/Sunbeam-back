package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class CompleteExcel implements Serializable {
	private String name = "新建excel";
	private List<SpreadSheet> SpreadSheet = new ArrayList<SpreadSheet>();

	public List<SpreadSheet> getSpreadSheet() {
		return SpreadSheet;
	}

	public void setSpreadSheet(List<SpreadSheet> spreadSheet) {
		SpreadSheet = spreadSheet;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
}
