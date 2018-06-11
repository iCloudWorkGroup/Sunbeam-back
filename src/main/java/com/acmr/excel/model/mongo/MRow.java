package com.acmr.excel.model.mongo;

import java.io.Serializable;

import com.acmr.excel.model.complete.OperProp;

public class MRow implements Serializable{
	
	private String sheetId;
	/*行别名*/
	private String alias;
	/*是否隐藏*/
	private Boolean hidden;
	/*序号*/
	private Integer sort;
	/*到顶部的高度*/
	private Integer top;
	/*高度*/
	private Integer height;
	/**/
	private OperProp props = new OperProp();
	
	
	public String getSheetId() {
		return sheetId;
	}
	public void setSheetId(String sheetId) {
		this.sheetId = sheetId;
	}
	public String getAlias() {
		return alias;
	}
	public void setAlias(String alias) {
		this.alias = alias;
	}
	public Boolean getHidden() {
		return hidden;
	}
	public void setHidden(Boolean hidden) {
		this.hidden = hidden;
	}
	public Integer getSort() {
		return sort;
	}
	public void setSort(Integer sort) {
		this.sort = sort;
	}
	public Integer getTop() {
		return top;
	}
	public void setTop(Integer top) {
		this.top = top;
	}
	public Integer getHeight() {
		return height;
	}
	public void setHeight(Integer height) {
		this.height = height;
	}
	public OperProp getProps() {
		return props;
	}
	public void setProps(OperProp props) {
		this.props = props;
	} 

}
