package com.acmr.excel.dao;

import com.acmr.excel.model.mongo.MSheet;

public interface MSheetDao {
 /**
  * 更新步骤
  * @param excelId
  * @param step
  * @param sheetId
  */
  void 	updateStep(String excelId,String sheetId,Integer step);
  /**
   * 更新冻结属性
   * @param mexcel
   * @param excelId
   */
  void updateFrozen(MSheet msheet,String excelId);
  
  /**
   * 解除冻结
   * @param mexcel
   * @param excelId
   */
  void updateUnFrozen(MSheet msheet,String excelId);
  
  /**
   * 更新最大列值
   * @param alias
   * @param excelId
   * @param sheetId
   */
  void updateMaxRow(Integer alias,String excelId,String sheetId);
  
  /**
   * 修改最大列值
   * @param alias
   * @param excelId
   * @param sheetId
   */
  void updateMaxCol(Integer alias,String excelId,String sheetId);
  /**
   * 根据sheetId获取MSheet对象
   * @param excelId
   * @param sheetId
   * @return
   */
  MSheet getMSheet(String excelId,String sheetId);
}
