package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MRowColCell;

public interface MRowColCellDao {

	
	/**
	 * 根据row号或列号，查询对应的关系映射表
	 * 
	 * @param excelId
	 *            数据集
	 * @param sheetId
	 * @param alias
	 *            行或列别名
	 * @param type
	 *            行或列类型
	 * @return
	 */
	List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			String alias, String type);

	/**
	 * 根据row号或列号list，查询对应的关系映射表
	 * 
	 * @param excelId
	 *            数据集
	 * @param sheetId
	 * @param aliasList
	 *            行或列别名集合
	 * @param type
	 *            行或列类型
	 * @return
	 */
	List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			List<String> aliasList, String type);

	/**
	 * 根据行、列id集合查询关系映射表
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param rowList
	 * @param colList
	 * @return
	 */
	List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			List<String> rowList, List<String> colList);
	
	/**
	 * 根据单元格id查询关系表
	 * @param excelId
	 * @param sheetId
	 * @param cellIdList
	 * @return
	 */
	List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			List<String> cellIdList);
	
	/**
	 * 根据cellId查找关系表
	 * @param excelId
	 * @param sheetId
	 * @param cellId
	 * @return
	 */
	List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			String cellId);

	/**
	 * 根据行别名删除关系表
	 * 
	 * @param excelId
	 *            数据集
	 * @param sheetId
	 * @param field
	 *            字段名 row或col
	 * @param alais
	 *            行或列的别名
	 */
	void delMRowColCell(String excelId, String sheetId, String field,
			String alias);

	/**
	 * 根据行、列别名集合删除对应关系表
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param rowList
	 * @param colList
	 */
	void delMRowColCellList(String excelId, String sheetId,
			List<String> rowList, List<String> colList);

	/**
	 * 根据单元格id集合，删除对应的关系表
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param cellIdList
	 */
	void delMRowColCellList1(String excelId, String sheetId,
			List<String> cellIdList);

	/**
	 * 批量根据cellId修改关系表的cellId
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param originalId
	 * @param modifiedId
	 */
	void updateMRowColCell(String excelId, String sheetId, String originalId,
			String modifiedId);

}
