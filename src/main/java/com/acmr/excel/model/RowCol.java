package com.acmr.excel.model;

import java.io.Serializable;

public class RowCol implements Serializable {
	private String alias;
	private Integer length;
	private Integer top;
	private String preAlias;// 用于指向它的前一行或前一列

	public String getAlias() {
		return alias;
	}

	public void setAlias(String alias) {
		this.alias = alias;
	}

	public Integer getLength() {
		return length;
	}

	public void setLength(Integer length) {
		this.length = length;
	}

	public Integer getTop() {
		return top;
	}

	public void setTop(Integer top) {
		this.top = top;
	}

	public String getPreAlias() {
		return preAlias;
	}

	public void setPreAlias(String preAlias) {
		this.preAlias = preAlias;
	}

}
