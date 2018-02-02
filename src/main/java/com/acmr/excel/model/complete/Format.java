package com.acmr.excel.model.complete;

public class Format {
	private String type;
	private Integer decimal;
	private Boolean thousands;
	private String dateFormat;
	private String currencySign;
	private String comment;
	private Boolean isValid = true;
	private Boolean currencyValid;

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public Integer getDecimal() {
		return decimal;
	}

	public void setDecimal(Integer decimal) {
		this.decimal = decimal;
	}

	public Boolean getThousands() {
		return thousands;
	}

	public void setThousands(Boolean thousands) {
		this.thousands = thousands;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}

	public String getCurrencySign() {
		return currencySign;
	}

	public void setCurrencySign(String currencySign) {
		this.currencySign = currencySign;
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

	public Boolean getCurrencyValid() {
		return currencyValid;
	}

	public void setCurrencyValid(Boolean currencyValid) {
		this.currencyValid = currencyValid;
	}

}
