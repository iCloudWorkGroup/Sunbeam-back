package com.acmr.excel.model.complete;

import java.io.Serializable;

public class BaseCell implements Serializable {
	private Content content = new Content();
	private Border border = new Border();
	private CustomProp customProp = new CustomProp();
	private Format format = new Format();
	
	public Format getFormat() {
		return format;
	}

	public void setFormat(Format format) {
		this.format = format;
	}

	public Content getContent() {
		return content;
	}

	public void setContent(Content content) {
		this.content = content;
	}

	public Border getBorder() {
		return border;
	}

	public void setBorder(Border border) {
		this.border = border;
	}

	public CustomProp getCustomProp() {
		return customProp;
	}

	public void setCustomProp(CustomProp customProp) {
		this.customProp = customProp;
	}

}
