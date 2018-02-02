package com.acmr.excel.service;

public interface StoreService {
	public Object get(String id);

	public boolean set(String id, Object object);
	
	public boolean update(String id, Object object,String express);
	
	public int getStep(String id ,Object object);
	
}
