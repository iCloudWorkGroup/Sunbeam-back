package com.acmr.excel.service;

import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.RowOperate;

public interface MRowService {

	/**
	 * 插入一行
	 * 
	 * @param rowOperate
	 * @param excelId
	 * @param step
	 *            步骤
	 */
	void insertRow(String excelId,RowOperate rowOperate,  Integer step);
	
	/**
	 * 批量增加行
	 * @param num
	 * @param excelId
	 * @param step
	 */
	void addRow(String excelId,int num,Integer step);

	/**
	 * 删除一行
	 * 
	 * @param rowOperate
	 * @param excelId
	 * @param step
	 *            步骤
	 */
	void delRow(String excelId,RowOperate rowOperate,  Integer step);

	/**
	 * 隐藏行
	 * 
	 * @param rowOperate
	 * @param excelId
	 * @param step
	 *            步骤
	 */
	void hideRow(String excelId,RowOperate rowOperate,  Integer step);

	/**
	 * 取消隐藏
	 * 
	 * @param rowOperate
	 * @param excelId
	 * @param step
	 *            步骤
	 */
	void showRow(String excelId,RowOperate rowOperate,  Integer step);

	/**
	 * 更新行长度
	 * 
	 * @param rowHeight
	 * @param excelId
	 * @param step
	 */
	void updateRowHeight(String excelId,RowHeight rowHeight,  Integer step);
}
