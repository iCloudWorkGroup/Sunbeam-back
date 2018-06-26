package com.acmr.excel.model.copy;

import java.io.Serializable;

public class Copy implements Serializable {
	private Orignal orignal;
	private Target target;

	public Orignal getOrignal() {
		return orignal;
	}

	public void setOrignal(Orignal orignal) {
		this.orignal = orignal;
	}

	public Target getTarget() {
		return target;
	}

	public void setTarget(Target target) {
		this.target = target;
	}

}
