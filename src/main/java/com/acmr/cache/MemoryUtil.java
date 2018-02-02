package com.acmr.cache;

import java.util.HashMap;
import java.util.Map;

public class MemoryUtil {
	private static Map<String,Object> map = new HashMap<String,Object>();

	public static Map<String, Object> getMap() {
		return map;
	}

	public static void setMap(Map<String, Object> map) {
		MemoryUtil.map = map;
	}
	
}
