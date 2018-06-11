package com.acmr.excel.model.mongo;

import java.util.ArrayList;
import java.util.List;

import com.acmr.excel.model.RowCol;

public class MRowColList {
	/*id用于区分是行或类 ，rList或cList*/
	private String id;
	private String sheetId;
	private List<RowCol> rcList = new ArrayList<RowCol>();
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	
	public String getSheetId() {
		return sheetId;
	}
	public void setSheetId(String sheetId) {
		this.sheetId = sheetId;
	}
	public List<RowCol> getRcList() {
		return rcList;
	}
	public void setRcList(List<RowCol> rcList) {
		this.rcList = rcList;
	}
	
	

}
