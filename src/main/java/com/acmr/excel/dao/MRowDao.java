package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MExcelRow;
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
	 * 取出所有的行样式
	 * @param excelId
	 * @param sheetId
	 * @return
	 */
	List<MRow> getMRowList(String excelId,String sheetId);
	
	/**
	 * 根据行别名，删除行样式
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 */
	void delMRow(String excelId,String sheetId,String alias);
	
	/**
	 * 根据别名，修改行隐藏属性
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 * @param status
	 */
	void updateRowHidden(String excelId,String sheetId,String alias,boolean status);
	
	/**
	 * 根据别名，返回一个行样式
	 * @param excelId
	 * @param sheetId
	 * @param alias
	 * @return
	 */
	MRow getMRow(String excelId,String sheetId,String alias);
	
	/**
	 * 根据行别名，更新行样式高度
	 * @param excelId
	 * @param alias
	 * @param height
	 */
	void updateRowHeight(String excelId,String sheetId,String alias,Integer height);
	
	/**
	 * 根据行别名数据集，修改对应的content属性
	 * @param property
	 * @param content
	 * @param aliasList
	 * @param excelId
	 * @param sheetId
	 */
	void updateContent(String property,Object content,List<String> aliasList,String excelId,String sheetId);

}
