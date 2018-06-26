package com.acmr.excel.model.complete;

import java.io.Serializable;

public class OperProp implements Serializable {
	private Border border = new Border();
	private Content content = new Content();
	private CustomProp customProp = new CustomProp();

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
