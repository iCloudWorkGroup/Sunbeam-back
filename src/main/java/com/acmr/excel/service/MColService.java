package com.acmr.excel.service;

import com.acmr.excel.model.ColOperate;
import com.acmr.excel.model.ColWidth;

public interface MColService {

	/**
	 * 插入一列
	 * 
	 * @param colOperate
	 * @param excelId
	 *            数据集
	 * @param step
	 *            步骤
	 */
	void insertCol(String excelId,ColOperate colOperate,  Integer step);
	
	void insertColDis(String excelId,ColOperate colOperate,  Integer step);
	
	/**
	 * 批量增加列
	 * @param num
	 * @param excelId
	 * @param step
	 */
	void addCol(String excelId,int num,Integer step);

	/**
	 * 删除一列
	 * 
	 * @param colOperate
	 * @param excelId
	 * @param step
	 *            步骤
	 */
	void delCol(String excelId,ColOperate colOperate,  Integer step);

	/**
	 * 隐藏列
	 * 
	 * @param colOperate
	 * @param excelId
	 * @param step
	 *            步骤
	 */
	void hideCol(String excelId,ColOperate colOperate,  Integer step);

	/**
	 * 取消隐藏
	 * 
	 * @param colOperate
	 * @param excelId、
	 * @param step
	 *            步骤
	 */
	void showCol(String excelId,ColOperate colOperate,  Integer step);

	/**
	 * 更新列宽度
	 * 
	 * @param colWidth
	 * @param excelId
	 * @param step
	 */
	void updateColWidth(String excelId,ColWidth colWidth,  Integer step);

}
