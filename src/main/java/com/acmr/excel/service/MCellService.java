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
	void saveContent(CellContent cell, int step, String excelId);

	/**
	 * 更新字体
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontFamily(Cell cell, int step, String excelId);

	/**
	 * 更新字体大小
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontSize(Cell cell, int step, String excelId);

	/**
	 * 更新字体weight
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontWeight(Cell cell, int step, String excelId);

	/**
	 * 更新字体是否斜体
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontItalic(Cell cell, int step, String excelId);

	/**
	 * 更新字体颜色
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontColor(Cell cell, int step, String excelId);

	/**
	 * 更新字体是否自动换行
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateWordwrap(Cell cell, int step, String excelId);

	/**
	 * 更新背景颜色
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateBgColor(Cell cell, int step, String excelId);

	/**
	 * 更新水平对齐方向
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateAlignlevel(Cell cell, int step, String excelId);

	/**
	 * 更新垂直对齐方向
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateAlignvertical(Cell cell, int step, String excelId);
	
	/**
	 * 更新字体下滑线状况
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateFontUnderline(Cell cell,int step,String excelId);

	/**
	 * 合并单元格
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void mergeCell(Cell cell, int step, String excelId);

	/**
	 * 拆分单元格
	 * 
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void splitCell(Cell cell, int step, String excelId);
	
	/**
	 * 设置边框样式
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateBorder(Cell cell,int step,String excelId);
	
	/**
	 * 设置单元格样式
	 * @param cellFormate
	 * @param step
	 * @param excelId
	 */
	void updateFormat(CellFormate cellFormate,int step,String excelId);
	
	/**
	 * 设置批注
	 * @param cellFormate
	 * @param step
	 * @param excelId
	 */
	void updateComment(Cell cell,int step,String excelId);
	
	/**
	 * 设置锁定状态
	 * @param cell
	 * @param step
	 * @param excelId
	 */
	void updateLock(Cell cell,int step,String excelId);

}
