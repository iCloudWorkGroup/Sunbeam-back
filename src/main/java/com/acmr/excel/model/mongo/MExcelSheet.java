package com.acmr.excel.model.mongo;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelSheetFreeze;
import acmr.util.ListHashMap;

public class MExcelSheet implements Serializable{
	
	private String name; // sheet名称
	private int hiddenstate; // 隐藏类型
	private ListHashMap<ExcelColumn> cols; // sheet的列集合

	private List<MExcelRowC> rows;// sheet的行集合

	private ExcelSheetFreeze freeze; // 是否有冻结区或者分隔区

	private Map<String, String> exps; // 扩展属性

	private int maxrow;
	private int maxcol;
	
	/**
	 * 构造函数
	 */
	public MExcelSheet() {
		cols = new ListHashMap<ExcelColumn>();
		rows = new ArrayList<MExcelRowC>();
		exps = new HashMap<String, String>();
		name = "new sheet";
		freeze = null;
		maxrow = 0;
		maxcol = 0;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getHiddenstate() {
		return hiddenstate;
	}

	public void setHiddenstate(int hiddenstate) {
		this.hiddenstate = hiddenstate;
	}

	public ListHashMap<ExcelColumn> getCols() {
		return cols;
	}

	public void setCols(ListHashMap<ExcelColumn> cols) {
		this.cols = cols;
	}

	public List<MExcelRowC> getRows() {
		return rows;
	}

	public void setRows(List<MExcelRowC> rows) {
		this.rows = rows;
	}

	public ExcelSheetFreeze getFreeze() {
		return freeze;
	}

	public void setFreeze(ExcelSheetFreeze freeze) {
		this.freeze = freeze;
	}

	public Map<String, String> getExps() {
		return exps;
	}

	public void setExps(Map<String, String> exps) {
		this.exps = exps;
	}

	public int getMaxrow() {
		return maxrow;
	}

	public void setMaxrow(int maxrow) {
		this.maxrow = maxrow;
	}

	public int getMaxcol() {
		return maxcol;
	}

	public void setMaxcol(int maxcol) {
		this.maxcol = maxcol;
	}

}
