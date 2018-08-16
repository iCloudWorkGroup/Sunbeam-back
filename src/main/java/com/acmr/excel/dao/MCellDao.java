package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MRowColCell;

public interface MCellDao {

	

	/**
	 * 根据cellID 查询cell列表
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param cellIdList
	 * @return
	 */
	List<MCell> getMCellList(String excelId, String sheetId,
			List<String> cellIdList);

	/**
	 * 根据MCell的id，返回一个MCell
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param id
	 * @return
	 */
	MCell getMCell(String excelId, String sheetId, String id);

	/**
	 * 根据cellId删除对应MExcelCell对象
	 * 
	 * @param excelId
	 * @param sheetId
	 * @param cellIdList
	 */
	void delMCell(String excelId, String sheetId, List<String> cellIdList);


	/**
	 * 根据cell id集合，更新MCell对象中，content中内容
	 * 
	 * @param property1
	 * @param value1
	 * @param property2
	 * @param value2
	 * @param idList
	 * @param excelId
	 * @param sheetId
	 */
	void updateContent(String property1, Object value1, String property2,
			Object value2, List<String> idList, String excelId, String sheetId);

	

	/**
	 * 更新单元格的边框
	 * 
	 * @param property
	 * @param value
	 * @param idList
	 * @param excelId
	 * @param sheetId
	 */
	void updateBorder(String property, Object value, List<String> idList,
			String excelId, String sheetId);

	/**
	 * 更新对应属性
	 * @param property
	 * @param value
	 * @param idList
	 * @param excelId
	 * @param sheetId
	 */
	void updateCustomProp(String property, Object value, List<String> idList,
			String excelId, String sheetId);

}
