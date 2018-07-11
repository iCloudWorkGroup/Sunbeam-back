package com.acmr.excel.distribute;

public class Target {
	/*类对象*/
	private Object object;
	/*方法名*/
	private String method;
	
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
	
	public Target(Object object,String method){
		this.object = object;
		this.method = method;
	}
	
}
