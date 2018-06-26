package com.acmr.excel.model.complete;

import java.io.Serializable;

public class CustomProp implements Serializable {

	private Boolean highlight;
	private String comment;

	public String getComment() {
		return comment;
	}

	public void setComment(String comment) {
		this.comment = comment;
	}

	public Boolean getHighlight() {
		return highlight;
	}

	public void setHighlight(Boolean highlight) {
		this.highlight = highlight;
	}

}
