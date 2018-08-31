package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MRow;

public interface MRowDao {

	/**
	 * 根据行别名list，获取行样式
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param rowList
	 * @return
	 */
	List<MRow> getMRowList(String excelId, String sheetId,
			List<String> rowList);

	/**
	 * 取出所有的行样式
	 * 
	 * @param excelId
	 * @param sheetId
	 * @return
	 */
	List<MRow> getMRowList(String excelId, String sheetId);

	/**
	 * 根据行别名，删除行样式
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 */
	void delMRow(String excelId, String sheetId, String alias);
	
	/**
	 * 根据行别名列表，删除行样式
	 * @param excelId
	 * @param sheetId
	 * @param aliasList
	 */
	void delMRowList(String excelId, String sheetId, List<String> aliasList);

	/**
	 * 根据别名，修改行隐藏属性
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 * @param status
	 */
	void updateRowHidden(String excelId, String sheetId, String alias,
			Boolean status);

	/**
	 * 根据别名，返回一个行样式
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 * @return
	 */
	MRow getMRow(String excelId, String sheetId, String alias);

	/**
	 * 根据行别名，更新行样式高度
	 * 
	 * @param excelId
	 * @param alias
	 * @param height
	 */
	void updateRowHeight(String excelId, String sheetId, String alias,
			Integer height);

	/**
	 * 根据行别名数据集，修改对应的content属性
	 * @param property1
	 * @param value1
	 * @param property2
	 * @param value2
	 * @param aliasList
	 * @param excelId
	 * @param sheetId
	 */
	void updateContent(String property1, Object value1, String property2,
			Object value2, List<String> aliasList, String excelId,
			String sheetId);

	/**
	 * 更新边框属性
	 * 
	 * @param property
	 * @param value
	 * @param aliasList
	 * @param excelId
	 * @param sheetId
	 */
	void updateBorderList(String property, Object value, List<String> aliasList,
			String excelId, String sheetId);

	/**
	 * 根据row别名，更新边框属性
	 * 
	 * @param property
	 * @param value
	 * @param alias
	 * @param excelId
	 * @param sheetId
	 */
	void updateBorder(String property, Object value, String alias,
			String excelId, String sheetId);
}
