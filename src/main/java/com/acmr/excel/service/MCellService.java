package com.acmr.excel.service;

import com.acmr.excel.model.Cell;


public interface MCellService {
	
	/**
	 * 保存单元格内容
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void saveContent(Cell cell,int step,String excelId);
	
	void updateFontFamily(Cell cell,int step,String excelId);

}
