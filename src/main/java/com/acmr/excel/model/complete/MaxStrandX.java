package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.Map;

public class MaxStrandX implements Serializable{
	private Map<String, Integer> maxAliasX;

	public Map<String, Integer> getMaxAliasX() {
		return maxAliasX;
	}

	public void setMaxAliasX(Map<String, Integer> maxAliasX) {
		this.maxAliasX = maxAliasX;
	}

}
