package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MExcelRow;

public interface MExcelRowDao {

/**
 * 根据行样式中code删除行样式
 * @param excelId
 * @param alais
 */
void delExcelRow(String excelId,String alias);

/**
 * 根据行样式中code，修改行隐藏属性值
 * @param excelId
 * @param alias
 * @param status
 */
void updateRowHidden(String excelId,String alias,boolean status);

/**
 * 根据行样式中code，跟新行高度
 * @param excelId
 * @param alias
 * @param height
 */
void updateRowHeight(String excelId,String alias,Integer height);

/**
 *根据code取到一个行对象
 * @param excelId
 * @param alias
 * @return
 */
MExcelRow getMExcelRow(String excelId,String alias);

/**
 * 更新字体属性
 * @param name
 * @param property
 * @param idList
 * @param excelId
 */
void updateFont(String name,Object property,String alias,String excelId);
}
