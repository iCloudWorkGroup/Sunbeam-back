package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MSheet;

public interface MSheetDao {
	/**
	 * 更新步骤
	 * 
	 * @param excelId
	 * @param step
	 * @param sheetId
	 */
	void updateStep(String excelId, String sheetId, Integer step);

	/**
	 * 更新冻结属性
	 * 
	 * @param msheet
	 * @param excelId
	 * @param sheetId
	 */
	void updateFrozen(MSheet msheet, String excelId, String sheetId);

	/**
	 * 解除冻结
	 * 
	 * @param msheet
	 * @param excelId
	 * @param sheetId
	 */
	void updateUnFrozen(MSheet msheet, String excelId, String sheetId);

	/**
	 * 更新最大列值
	 * 
	 * @param alias
	 * @param excelId
	 * @param sheetId
	 */
	void updateMaxRow(Integer alias, String excelId, String sheetId);

	/**
	 * 修改最大列值
	 * 
	 * @param alias
	 * @param excelId
	 * @param sheetId
	 */
	void updateMaxCol(Integer alias, String excelId, String sheetId);

	/**
	 * 根据sheetId获取MSheet对象
	 * 
	 * @param excelId
	 * @param sheetId
	 * @return
	 */
	MSheet getMSheet(String excelId, String sheetId);
	
	/**
	 * 根据属性名更新，msheet对应属性
	 * @param excelId
	 * @param sheetId
	 * @param proterty 属性名
	 * @param value    属性值
	 */
	void updateMSheetProperty(String excelId,String sheetId,String proterty,Object value);
	
	/**
	 * 获取所有的表（数据集）
	 * @return
	 */
	List<String> getTableList();
	
	/**
	 * 根据传过来的msheet对象更新对应的属性
	 * @param excelId
	 * @param sheetId
	 * @param msheet
	 */
	void updateMSheetByObject(String excelId,String sheetId,MSheet msheet);
}
