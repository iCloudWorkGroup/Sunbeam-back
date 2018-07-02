package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MCol;

public interface MColDao {

	/**
	 * 根据列别名list，获取列样式表
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param colList
	 * @return
	 */
	List<MCol> getMColList(String excelId, String sheetId,
			List<String> colList);

	/**
	 * 获取所有的列样式
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param colList
	 * @return
	 */
	List<MCol> getMColList(String excelId, String sheetId);

	/**
	 * 根据列别名，删除列样式表
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 */
	void delExcelCol(String excelId, String sheetId, String alias);

	/**
	 * 根据列样式中的别名值，修改列样式的隐藏属性
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 * @param status
	 */
	void updateColHiddenStatus(String excelId, String sheetId, String alias,
			boolean status);

	/**
	 * 根据列别名，返回一个列样式对象
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 * @return
	 */
	MCol getMCol(String excelId, String sheetId, String alias);

	/**
	 * 根据列别名，更新列样式宽度
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 * @param width
	 */
	void updateColWidth(String excelId, String sheetId, String alias,
			Integer width);

	/**
	 * 根据列别名list，更新列content
	 * 
	 * @param property
	 * @param content
	 * @param aliasList
	 * @param excelId
	 * @param sheetId
	 */
	void updateContent(String property, Object content, List<String> aliasList,
			String excelId, String sheetId);
	
	/**
	 * 更新边框属性
	 * @param property
	 * @param value
	 * @param aliasList
	 * @param excelId
	 * @param sheetId
	 */
	void updateBorder(String property, Object value, List<String> aliasList,
			String excelId, String sheetId);
    /**
     * 根据col别名，更新边框属性
     * @param property
     * @param value
     * @param alias
     * @param excelId
     * @param sheetId
     */
	void updateBorder(String property, Object value, String alias,
			String excelId, String sheetId);
}
