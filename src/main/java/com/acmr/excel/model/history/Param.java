package com.acmr.excel.model.history;

import java.io.Serializable;

public class Param implements Serializable{
	//实例对象
	private String target;
	//方法名称
	private String name;
	//参数列表
	private Object[] arguments;

	public String getTarget() {
		return target;
	}

	public void setTarget(String target) {
		this.target = target;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Object[] getArguments() {
		return arguments;
	}

	public void setArguments(Object[] arguments) {
		this.arguments = arguments;
	}
	
	public Param(String target,String name,Object[] arguments){
		this.target = target;
		this.name=name;
		this.arguments =arguments; 
	}
	
	public Param(){}

}
