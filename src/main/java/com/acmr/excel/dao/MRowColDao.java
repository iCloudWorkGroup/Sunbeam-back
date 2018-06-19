package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.RowCol;


public interface MRowColDao {
	
	/**
	 * 根据excelId，查询出排序之后的行列表
	 * @param sortRcList 排序之后的行集合
	 * @param excelId
	 */
	void getRowList(List<RowCol> sortRcList,String  excelId,String sheetId);
	
	/**
	 * 根据excelId，查询出排序之后的列列表
	 * @param sortClList 排序之后的列集合
	 * @param excelId
	 */
	void getColList(List<RowCol> sortClList,String excelId,String sheetId);
	
	/**
	 * 根据条件修改行或列前一个行或列的别名
	 * @param excelId  表ID
	 * @param sheetId
	 * @param id     RowColList ID rList或cList
	 * @param alias   行或者列别名
	 * @param preAlias   前面一行或者列别名
	 */
	void updateRowCol(String excelId,String sheetId, String id,String alias,String preAlias);
	
	/**
	 * 插入一条行或列记录
	 * @param excelId  数据集ID
	 * @param sheetId  sheet ID
	 * @param rowCol   行列对象
	 * @param id       行或列集合ID
	 */
	void insertRowCol(String excelId,String sheetId,RowCol rowCol,String id);

	/**
	 * 根据条件删除一条行或类记录
	 * @param excelId  数据集ID
	 * @param sheetId  
	 * @param alias   别名
	 * @param id       行或列集合ID
	 */
    void delRowCol(String excelId,String sheetId,String alias,String id);
    
    /**
     * 根据别名更新行或列长度
     * @param excelId
     * @param sheetId
     * @param id
     * @param alias
     * @param length
     */
    void updateRowColLength(String excelId,String sheetId,String id,String alias,Integer length);
}
