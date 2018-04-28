package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.Map;

public class MaxStrandY implements Serializable{
	private Map<String, Integer> maxAliasY;

	public Map<String, Integer> getMaxAliasY() {
		return maxAliasY;
	}

	public void setMaxAliasY(Map<String, Integer> maxAliasY) {
		this.maxAliasY = maxAliasY;
	}
}
