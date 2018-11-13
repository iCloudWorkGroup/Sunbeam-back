package com.acmr.excel.service;

import com.acmr.excel.model.Cell;
import com.acmr.excel.model.CellContent;
import com.acmr.excel.model.CellFormate;

public interface MCellService {

	/**
	 * 保存单元格内容
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void saveContent(String excelId,CellContent cell, int step);
	
	/**
	 * 清除单元格
	 * @param excelId
	 * @param cell
	 * @param step
	 */
	void clearContent(String excelId,Cell cell, int step);

	/**
	 * 更新字体
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontFamily(String excelId,Cell cell, int step);

	/**
	 * 更新字体大小
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontSize(String excelId,Cell cell, int step);

	/**
	 * 更新字体weight
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontWeight(String excelId,Cell cell, int step);

	/**
	 * 更新字体是否斜体
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontItalic(String excelId,Cell cell, int step);

	/**
	 * 更新字体颜色
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontColor(String excelId,Cell cell, int step);

	/**
	 * 更新字体是否自动换行
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateWordwrap(String excelId,Cell cell, int step);

	/**
	 * 更新背景颜色
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateBgColor(String excelId,Cell cell, int step);

	/**
	 * 更新水平对齐方向
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateAlignlevel(String excelId,Cell cell, int step);

	/**
	 * 更新垂直对齐方向
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateAlignvertical(String excelId,Cell cell, int step);
	
	/**
	 * 更新字体下滑线状况
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontUnderline(String excelId,Cell cell,int step);

	/**
	 * 合并单元格
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void mergeCell(String excelId,Cell cell, int step);

	/**
	 * 拆分单元格
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void splitCell(String excelId,Cell cell, int step);
	
	/**
	 * 设置边框样式
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateBorder(String excelId,Cell cell,int step);
	
	/**
	 * 设置单元格样式
	 * @param cellFormate
	 * @param step
	 * @param excelId
	 */
	void updateFormat(String excelId,CellFormate cellFormate,int step);
	
	/**
	 * 设置批注
	 * @param cellFormate
	 * @param step
	 * @param excelId
	 */
	void updateComment(String excelId,Cell cell,int step);
	
	/**
	 * 设置锁定状态
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateLock(String excelId,Cell cell,int step);

}
