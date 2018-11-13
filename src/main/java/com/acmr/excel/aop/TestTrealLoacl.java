package com.acmr.excel.aop;

import java.util.List;

public class TestTrealLoacl {
	private ThreadLocal<String> excelId;
	private ThreadLocal<List<Object>> list;
	
	
	public void a(){
		excelId = new ThreadLocal<String>();
		list = new ThreadLocal<List<Object>>();
	}
	
	public void b(){
		String a = excelId.get();
		List<Object> li = list.get();
	}
}
