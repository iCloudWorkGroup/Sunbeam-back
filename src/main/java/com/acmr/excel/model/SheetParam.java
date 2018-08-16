package com.acmr.excel.model;

import java.io.Serializable;

public class SheetParam implements Serializable{

	private boolean protect;
	
	private String passwd;

	public boolean isProtect() {
		return protect;
	}

	public void setProtect(boolean protect) {
		this.protect = protect;
	}

	public String getPasswd() {
		return passwd;
	}

	public void setPasswd(String passwd) {
		this.passwd = passwd;
	}
	
}
