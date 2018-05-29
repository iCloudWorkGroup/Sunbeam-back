package com.acmr.excel.model.complete;

import java.io.Serializable;

/***
 * 返回到页面的列属性
 * @author liucb
 *
 */
public class Glx implements Serializable {
	private String alias;
	private int left;
	private int width;
	private Boolean hidden;
	private int sort;
	
	private OperProp props = new OperProp();

    
	public String getAlias() {
		return alias;
	}

	public void setAlias(String alias) {
		this.alias = alias;
	}

	public int getSort() {
		return sort;
	}

	public void setSort(int sort) {
		this.sort = sort;
	}

	public OperProp getProps() {
		return props;
	}

	public void setProps(OperProp props) {
		this.props = props;
	}

	public int getLeft() {
		return left;
	}

	public void setLeft(int left) {
		this.left = left;
	}

	public int getWidth() {
		return width;
	}

	public void setWidth(int width) {
		this.width = width;
	}

	public Boolean getHidden() {
		return hidden;
	}

	public void setHidden(Boolean hidden) {
		this.hidden = hidden;
	}

}
