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
	void insertCol(ColOperate colOperate, String excelId, Integer step);
	
	void insertColDis(ColOperate colOperate, String excelId, Integer step);
	
	/**
	 * 批量增加列
	 * @param num
	 * @param excelId
	 * @param step
	 */
	void addCol(int num,String excelId,Integer step);

	/**
	 * 删除一列
	 * 
	 * @param colOperate
	 * @param excelId
	 * @param step
	 *            步骤
	 */
	void delCol(ColOperate colOperate, String excelId, Integer step);

	/**
	 * 隐藏列
	 * 
	 * @param colOperate
	 * @param excelId
	 * @param step
	 *            步骤
	 */
	void hideCol(ColOperate colOperate, String excelId, Integer step);

	/**
	 * 取消隐藏
	 * 
	 * @param colOperate
	 * @param excelId、
	 * @param step
	 *            步骤
	 */
	void showCol(ColOperate colOperate, String excelId, Integer step);

	/**
	 * 更新列宽度
	 * 
	 * @param colWidth
	 * @param excelId
	 * @param step
	 */
	void updateColWidth(ColWidth colWidth, String excelId, Integer step);

}
