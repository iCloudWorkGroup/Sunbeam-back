package com.acmr.excel.model.complete;

import java.util.ArrayList;
import java.util.List;

public class Validate {

	private Rule rule;
	private List<Coordinate> coordinates = new ArrayList<Coordinate>();
	
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
	
}
