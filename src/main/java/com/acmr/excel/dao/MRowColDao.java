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
	
	/**
	 * 根据条件修改行或列前一个行或列的别名
	 * @param excelId  表ID
	 * @param id     RowColList ID rList或cList
	 * @param alias   行或者列别名
	 * @param preAlias   前面一行或者列别名
	 */
	void updateRowCol(String excelId, String id,String alias,String preAlias);
	
	/**
	 * 根据条件更新一条行或类记录
	 * @param excelId  数据集ID
	 * @param rowCol   行列对象
	 * @param id       行或列集合ID
	 */
	void insertRowCol(String excelId,RowCol rowCol,String id);

	/**
	 * 根据条件删除一条行或类记录
	 * @param excelId  数据集ID
	 * @param alias   别名
	 * @param id       行或列集合ID
	 */
    void delRowCol(String excelId,String alias,String id);
    
    void updateRowColLength(String excelId,String id,String alias,Integer length);
}
