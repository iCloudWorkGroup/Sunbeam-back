package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

public class StrandY implements Serializable{
	private Map<String,Map<String,Integer>> aliasY = new HashMap<String, Map<String,Integer>>();

	public Map<String, Map<String, Integer>> getAliasY() {
		return aliasY;
	}

	public void setAliasY(Map<String, Map<String, Integer>> aliasY) {
		this.aliasY = aliasY;
	}

	

	
}
