package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

public class VersionHistory implements Serializable{

	private Map<Integer, History> map = new HashMap<Integer, History>();
	private Map<Integer, Integer> version = new HashMap<Integer, Integer>();

	public Map<Integer, Integer> getVersion() {
		return version;
	}

	public void setVersion(Map<Integer, Integer> version) {
		this.version = version;
	}

	public Map<Integer, History> getMap() {
		return map;
	}

	public void setMap(Map<Integer, History> map) {
		this.map = map;
	}

}
