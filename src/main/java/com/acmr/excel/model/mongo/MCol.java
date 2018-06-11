package com.acmr.excel.model.mongo;

import java.io.Serializable;

import com.acmr.excel.model.complete.OperProp;

public class MCol implements Serializable{
	
    private String sheetId;
	/*列别名*/
	private String alias;
	/*和左边的距离*/
	private Integer left;
	/*宽度*/
	private Integer width;
	/*是否隐藏*/
	private Boolean hidden;
	/*序号*/
	private Integer sort;
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
	public Integer getLeft() {
		return left;
	}
	public void setLeft(Integer left) {
		this.left = left;
	}
	public Integer getWidth() {
		return width;
	}
	public void setWidth(Integer width) {
		this.width = width;
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
	public OperProp getProps() {
		return props;
	}
	public void setProps(OperProp props) {
		this.props = props;
	}
	
}
