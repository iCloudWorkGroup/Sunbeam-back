package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import com.acmr.excel.model.mongo.MCell;

public class MCellBefore implements Serializable{

	private List<String> idList = new ArrayList<String>();

	private List<MCell> mcellList = new ArrayList<MCell>();

	public List<String> getIdList() {
		return idList;
	}

	public void setIdList(List<String> idList) {
		this.idList = idList;
	}

	public List<MCell> getMcellList() {
		return mcellList;
	}

	public void setMcellList(List<MCell> mcellList) {
		this.mcellList = mcellList;
	}
	
	public MCellBefore(List<String> idList,List<MCell> mcellList){
		this.idList = idList;
		this.mcellList = mcellList;
	}
	
	public MCellBefore(){
		
	}

}
