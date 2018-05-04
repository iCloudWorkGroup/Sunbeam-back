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
	private int originWidth;
	private int sort;
	private boolean hidden;
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

	public int getOriginWidth() {
		return originWidth;
	}

	public void setOriginWidth(int originWidth) {
		this.originWidth = originWidth;
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


	public boolean isHidden() {
		return hidden;
	}

	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

}
