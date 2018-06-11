package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MRow;

public interface MRowDao {
	/**
	 * 根据行别名list，获取行样式
	 * @param excelId
	 * @param sheetId
	 * @param rowList
	 * @return
	 */
	List<MRow> getMRowList(String excelId,String sheetId,List<String> rowList);
	
	/**
	 * 根据行别名，删除行样式
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 */
	void delMRow(String excelId,String sheetId,String alias);

}
