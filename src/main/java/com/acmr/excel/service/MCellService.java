package com.acmr.excel.service;

import com.acmr.excel.model.Cell;


public interface MCellService {
	
	void saveContent(Cell cell,int step,String excelId);

}
