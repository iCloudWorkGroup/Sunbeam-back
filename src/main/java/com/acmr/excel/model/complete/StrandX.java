package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

public class StrandX implements Serializable{
	private Map<String,Map<String,Integer>> aliasX = new HashMap<String, Map<String,Integer>>();

	public Map<String, Map<String, Integer>> getAliasX() {
		return aliasX;
	}

	public void setAliasX(Map<String, Map<String, Integer>> aliasX) {
		this.aliasX = aliasX;
	}
	
}
