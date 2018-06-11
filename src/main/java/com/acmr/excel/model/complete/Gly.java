package com.acmr.excel.model.complete;

import java.io.Serializable;

import com.acmr.excel.model.mongo.MRow;

/***
 * 返回到页面的行属性
 * @author liucb
 *
 */
public class Gly implements Serializable {
	private String alias;
	private Boolean hidden;
	private int sort;
	private int top;
	private int height;
	
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

	public Boolean getHidden() {
		return hidden;
	}

	public void setHidden(Boolean hidden) {
		this.hidden = hidden;
	}
	
	public Gly(MRow mrow){
		this.alias = mrow.getAlias();
		this.height = mrow.getHeight();
		this.hidden = mrow.getHidden();
		this.props = mrow.getProps();
	}
	
	public Gly(){
		
	}

}
