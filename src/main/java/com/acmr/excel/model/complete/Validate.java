package com.acmr.excel.model.complete;

import java.util.ArrayList;
import java.util.List;

public class Validate {

	private Rule rule;
	private List<Coordinate> coordinates = new ArrayList<Coordinate>();
	private int count;//用于统计有多少个单元格引用这条验证

	public Rule getRule() {
		return rule;
	}

	public void setRule(Rule rule) {
		this.rule = rule;
	}

	public List<Coordinate> getCoordinates() {
		return coordinates;
	}

	public void setCoordinates(List<Coordinate> coordinates) {
		this.coordinates = coordinates;
	}

	public int getCount() {
		return count;
	}

	public void setCount(int count) {
		this.count = count;
	}

}
