package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class SheetElement implements Serializable{
	private List<Glx> glX = new ArrayList<Glx>();
	private List<Gly> glY = new ArrayList<Gly>();
	private Position posi = new Position();
	private List<OneCell> cells = new ArrayList<OneCell>();
	private List<PresetCell> presetCells = new ArrayList<PresetCell>();
	private Frozen frozen = new Frozen();
	private StartAlais startAlais = new StartAlais();

	public List<Glx> getGlX() {
		return glX;
	}

	public void setGlX(List<Glx> glX) {
		this.glX = glX;
	}

	public List<Gly> getGlY() {
		return glY;
	}

	public void setGlY(List<Gly> glY) {
		this.glY = glY;
	}

	public Position getPosi() {
		return posi;
	}

	public void setPosi(Position posi) {
		this.posi = posi;
	}

	public List<OneCell> getCells() {
		return cells;
	}

	public void setCells(List<OneCell> cells) {
		this.cells = cells;
	}

	public Frozen getFrozen() {
		return frozen;
	}

	public void setFrozen(Frozen frozen) {
		this.frozen = frozen;
	}

	public List<PresetCell> getPresetCells() {
		return presetCells;
	}

	public void setPresetCells(List<PresetCell> presetCells) {
		this.presetCells = presetCells;
	}

	public StartAlais getStartAlais() {
		return startAlais;
	}

	public void setStartAlais(StartAlais startAlais) {
		this.startAlais = startAlais;
	}

}
