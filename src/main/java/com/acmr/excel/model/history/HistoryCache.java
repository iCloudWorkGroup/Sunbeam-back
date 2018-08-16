package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class HistoryCache implements Serializable{
  
	//指针，用于标识执行到哪个步骤
	private int index;
	//用于标识上一个方法名称
	private String flag;
	
	private List<History> list = new ArrayList<History>();

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public String getFlag() {
		return flag;
	}

	public void setFlag(String flag) {
		this.flag = flag;
	}

	public List<History> getList() {
		return list;
	}

	public void setList(List<History> list) {
		this.list = list;
	}
	
    public HistoryCache(){}
    
    public HistoryCache(int index,List<History> list){
    	this.index = index;
    	this.list = list;
    }
    
    public HistoryCache(int index,String flag,List<History> list){
    	this.index = index;
    	this.flag = flag;
    	this.list = list;
    }
    
}
