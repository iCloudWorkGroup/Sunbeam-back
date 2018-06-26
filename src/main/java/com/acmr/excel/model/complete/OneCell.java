package com.acmr.excel.model.complete;

import java.io.Serializable;

import com.acmr.excel.model.mongo.MCell;

public class OneCell implements Serializable {
	private Border border = new Border();
	private Content content = new Content();
	private CustomProp customProp = new CustomProp();
	private Occupy occupy = new Occupy();

	public Border getBorder() {
		return border;
	}

	public void setBorder(Border border) {
		this.border = border;
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

	public Occupy getOccupy() {
		return occupy;
	}

	public void setOccupy(Occupy occupy) {
		this.occupy = occupy;
	}

	public OneCell(MCell mcell) {
		this.border = mcell.getBorder();
		this.content = mcell.getContent();
		this.customProp = mcell.getCustomProp();
	}

	public OneCell() {

	}

}
