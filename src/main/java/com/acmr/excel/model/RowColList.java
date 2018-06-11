package com.acmr.excel.model;

import java.util.ArrayList;
import java.util.List;

public class RowColList {
	private String id;
	private List<RowCol> rcList = new ArrayList<RowCol>();

	public List<RowCol> getRcList() {
		return rcList;
	}

	public void setRcList(List<RowCol> rcList) {
		this.rcList = rcList;
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

}
