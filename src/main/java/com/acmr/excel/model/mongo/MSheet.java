package com.acmr.excel.model.mongo;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import com.acmr.excel.model.complete.Rule;
import com.acmr.excel.model.complete.Validate;

public class MSheet implements Serializable {
	/* sheetId */
	private String id;
	/* 操作不住 */
	private Integer step;
	/* sheet名称 */
	private String sheetName;
	/* 最大行 */
	private Integer maxrow;
	/* 最大列 */
	private Integer maxcol;
	/* 是否冻结 */
	private Boolean freeze;
	/* 冻结列坐标点 */
	private String colAlias;
	/* 冻结行坐标点 */
	private String rowAlias;
	/* 可视列坐标点 */
	private String viewColAlias;
	/* 可视行坐标点 */
	private String viewRowAlias;
	/* sheet序号 */
	private Integer sort;
	/* 是否锁定 */
	private Boolean protect;
	
	private String  passwd;
	
	private List<Rule> rule = new ArrayList<Rule>();

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public Integer getStep() {
		return step;
	}

	public void setStep(Integer step) {
		this.step = step;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public Integer getMaxrow() {
		return maxrow;
	}

	public void setMaxrow(Integer maxrow) {
		this.maxrow = maxrow;
	}

	public Integer getMaxcol() {
		return maxcol;
	}

	public void setMaxcol(Integer maxcol) {
		this.maxcol = maxcol;
	}

	public Boolean getFreeze() {
		return freeze;
	}

	public void setFreeze(Boolean freeze) {
		this.freeze = freeze;
	}

	public String getColAlias() {
		return colAlias;
	}

	public void setColAlias(String colAlias) {
		this.colAlias = colAlias;
	}

	public String getRowAlias() {
		return rowAlias;
	}

	public void setRowAlias(String rowAlias) {
		this.rowAlias = rowAlias;
	}

	public String getViewColAlias() {
		return viewColAlias;
	}

	public void setViewColAlias(String viewColAlias) {
		this.viewColAlias = viewColAlias;
	}

	public String getViewRowAlias() {
		return viewRowAlias;
	}

	public void setViewRowAlias(String viewRowAlias) {
		this.viewRowAlias = viewRowAlias;
	}

	public Integer getSort() {
		return sort;
	}

	public void setSort(Integer sort) {
		this.sort = sort;
	}

	public Boolean getProtect() {
		return protect;
	}

	public void setProtect(Boolean protect) {
		this.protect = protect;
	}

	public String getPasswd() {
		return passwd;
	}

	public void setPasswd(String passwd) {
		this.passwd = passwd;
	}

	public List<Rule> getRule() {
		return rule;
	}

	public void setRule(List<Rule> rule) {
		this.rule = rule;
	}
	
}
