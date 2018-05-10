package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.position.RowCol;

public interface MRowColDao {
	
	/**
	 * 根据excelId，查询出排序之后的行列表
	 * @param sortRcList 排序之后的行集合
	 * @param excelId
	 */
	void getRowList(List<RowCol> sortRcList,String  excelId);
	
	/**
	 * 根据excelId，查询出排序之后的列列表
	 * @param sortClList 排序之后的列集合
	 * @param excelId
	 */
	void getColList(List<RowCol> sortClList,String excelId);


}
