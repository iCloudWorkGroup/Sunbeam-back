package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MCol;

public interface MColDao {
    
	/**
	 * 根据列别名list，获取列样式表
	 * @param excelId
	 * @param sheetId
	 * @param colList
	 * @return
	 */
	List<MCol> getMColList(String excelId,String sheetId,List<String> colList);
	
	/**
	 * 根据列别名，删除列样式表
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 */
	void delExcelCol(String excelId,String sheetId,String alias);
}
