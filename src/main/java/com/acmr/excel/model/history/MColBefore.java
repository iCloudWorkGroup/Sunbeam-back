package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import com.acmr.excel.model.mongo.MCol;

public class MColBefore implements Serializable{
  
	private List<String> idList = new ArrayList<String>();
	
	private List<MCol> mcolList = new ArrayList<MCol>();

	public List<String> getIdList() {
		return idList;
	}

	public void setIdList(List<String> idList) {
		this.idList = idList;
	}

	public List<MCol> getMcolList() {
		return mcolList;
	}

	public void setMcolList(List<MCol> mcolList) {
		this.mcolList = mcolList;
	}
	
	public MColBefore(List<String> idList,List<MCol> mcolList){
		this.idList = idList;
		this.mcolList = mcolList;
	}
	public MColBefore(){
		
	}
}
