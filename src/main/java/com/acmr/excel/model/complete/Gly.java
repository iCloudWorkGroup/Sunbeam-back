package com.acmr.excel.model.complete;

import java.io.Serializable;

/***
 * 返回到页面的行属性
 * @author liucb
 *
 */
public class Gly implements Serializable {
	private String alias;
	private int top;
	private int height;
	private int originHeight;
	private int sort;
	private boolean hidden;
	private OperProp props = new OperProp();



	public int getTop() {
		return top;
	}

	public void setTop(int top) {
		this.top = top;
	}

	public int getHeight() {
		return height;
	}

	public void setHeight(int height) {
		this.height = height;
	}


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

	public boolean isHidden() {
		return hidden;
	}

	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	public int getOriginHeight() {
		return originHeight;
	}

	public void setOriginHeight(int originHeight) {
		this.originHeight = originHeight;
	}

}
