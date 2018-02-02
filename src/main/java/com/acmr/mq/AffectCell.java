package com.acmr.mq;

import java.io.Serializable;

public class AffectCell implements Serializable {
	private String startColAlias;
	private String endColAlias;
	private String startRowAlias;
	private String endRowAlias;
	private String type; // rows_insert rows_delete cols_insert cols_delete
							// cols_width rows_height

	public String getStartColAlias() {
		return startColAlias;
	}

	public void setStartColAlias(String startColAlias) {
		this.startColAlias = startColAlias;
	}

	public String getEndColAlias() {
		return endColAlias;
	}

	public void setEndColAlias(String endColAlias) {
		this.endColAlias = endColAlias;
	}

	public String getStartRowAlias() {
		return startRowAlias;
	}

	public void setStartRowAlias(String startRowAlias) {
		this.startRowAlias = startRowAlias;
	}

	public String getEndRowAlias() {
		return endRowAlias;
	}

	public void setEndRowAlias(String endRowAlias) {
		this.endRowAlias = endRowAlias;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

}
