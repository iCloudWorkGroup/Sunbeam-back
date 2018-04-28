package com.acmr.excel.model.complete;

import java.io.Serializable;

public class OneCell extends BaseCell implements Serializable {
	private Occupy occupy = new Occupy();
	private boolean wordWrap;
	private boolean highlight;
	private boolean hidden;

	public boolean isHidden() {
		return hidden;
	}

	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	public boolean isHighlight() {
		return highlight;
	}

	public void setHighlight(boolean highlight) {
		this.highlight = highlight;
	}

	public Occupy getOccupy() {
		return occupy;
	}

	public void setOccupy(Occupy occupy) {
		this.occupy = occupy;
	}

	public boolean isWordWrap() {
		return wordWrap;
	}

	public void setWordWrap(boolean wordWrap) {
		this.wordWrap = wordWrap;
	}
}
