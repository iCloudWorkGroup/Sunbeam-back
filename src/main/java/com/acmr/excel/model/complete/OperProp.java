package com.acmr.excel.model.complete;

import java.io.Serializable;


public class OperProp implements Serializable {
	private Content content = new Content();
	private Border border = new Border();
	private CustomProp customProp = new CustomProp();
	private Format formate = new Format();

	public Format getFormate() {
		return formate;
	}

	public void setFormate(Format formate) {
		this.formate = formate;
	}

	public Content getContent() {
		return content;
	}

	public void setContent(Content content) {
		this.content = content;
	}

	public CustomProp getCustomProp() {
		return customProp;
	}

	public void setCustomProp(CustomProp customProp) {
		this.customProp = customProp;
	}

	public Border getBorder() {
		return border;
	}

	public void setBorder(Border border) {
		this.border = border;
	}

}
