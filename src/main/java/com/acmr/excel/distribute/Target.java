package com.acmr.excel.distribute;

public class Target {
	/*类对象*/
	private Object object;
	/*方法名*/
	private String method;
	
	private int type;
	
	
	public Object getObject() {
		return object;
	}
	public void setObject(Object object) {
		this.object = object;
	}
	public String getMethod() {
		return method;
	}
	public void setMethod(String method) {
		this.method = method;
	}
	
	public int getType() {
		return type;
	}
	public void setType(int type) {
		this.type = type;
	}
	public Target(Object object,String method,int type){
		this.object = object;
		this.method = method;
		this.type = type;
	}
	
}
