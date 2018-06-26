package com.acmr.excel.model.mongo;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import acmr.excel.pojo.ExcelRow;

public class MExcelRowC implements Serializable {

	private int height; // 网页上行高 poi高度=网页高度*18 行高 实际高度*20

	private String code; // 唯一编号

	private List<MExcelCell> cells; // 行中的单元格集合

	private Map<String, String> exps; // 扩展属性

	private boolean rowhidden;

	private boolean inlist; // 是否已经放到队列中了，如果放在队列中了就不能再修改code

	public MExcelRowC() {
		cells = new ArrayList<MExcelCell>();
		exps = new HashMap<String, String>();
		code = "-1";
		// height = 270;
		height = 19;
		inlist = false;
		rowhidden = false;
	}

	/**
	 * 构造函数，默认高270 为13.5
	 */
	public MExcelRowC(String code1) {
		cells = new ArrayList<MExcelCell>();
		exps = new HashMap<String, String>();
		code = code1;
		// height = 270;
		height = 19;
		inlist = false;
		rowhidden = false;
	}

	public MExcelRowC(ExcelRow row) {
		cells = new ArrayList<MExcelCell>();
		exps = row.getExps();
		code = row.getCode();
		// height = 270;
		height = row.getHeight();
		inlist = false;
		rowhidden = false;
	}

	public int getHeight() {
		return height;
	}

	public void setHeight(int height) {
		this.height = height;
	}

	public String getCode() {
		return code;
	}

	public void setCode(String code) {
		this.code = code;
	}

	public List<MExcelCell> getCells() {
		return cells;
	}

	public void setCells(List<MExcelCell> cells) {
		this.cells = cells;
	}

	public Map<String, String> getExps() {
		return exps;
	}

	public void setExps(Map<String, String> exps) {
		this.exps = exps;
	}

	public boolean isRowhidden() {
		return rowhidden;
	}

	public void setRowhidden(boolean rowhidden) {
		this.rowhidden = rowhidden;
	}

	public boolean isInlist() {
		return inlist;
	}

	public void setInlist(boolean inlist) {
		this.inlist = inlist;
	}

}
