package com.acmr.excel.model.complete;

import java.io.Serializable;

public class CustomProp implements Serializable {
	private String background = "rgb(255,255,255)";
	private String bgRgbColor;
//	private String format = "normal";
//	private String remarket;
//	private Integer decimal;
//	private Boolean thousands;
//	private String dateFormat;
//	private String currencySign;
	private String comment;
	/**
	 * 文本内容，与设置类型是否匹配
	 */
	private Boolean isValid = true;

	public String getBackground() {
		return background;
	}

	public void setBackground(String background) {
		this.background = background;
	}

	public String getBgRgbColor() {
		return bgRgbColor;
	}

	public void setBgRgbColor(String bgRgbColor) {
		this.bgRgbColor = bgRgbColor;
	}


	public String getComment() {
		return comment;
	}

	public void setComment(String comment) {
		this.comment = comment;
	}

	public Boolean getIsValid() {
		return isValid;
	}

	public void setIsValid(Boolean isValid) {
		this.isValid = isValid;
	}

}
