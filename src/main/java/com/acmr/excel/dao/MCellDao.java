package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MRowColCell;

public interface MCellDao {

 /**
  * 根据row号或列号，查询对应的关系映射表
  * @param excelId  数据集
  * @param sheetId   
  * @param alias  行或列别名
  * @param type   行或列类型
  * @return 
  */
 List<MRowColCell> getMRowColCellList(String excelId,String sheetId,String alias,String type);
 
 /**
  * 根据row号或列号list，查询对应的关系映射表
  * @param excelId  数据集
  * @param sheetId   
  * @param aliasList  行或列别名集合
  * @param type   行或列类型
  * @return 
  */
 List<MRowColCell> getMRowColCellList(String excelId,String sheetId,List<String> aliasList,String type);
 
 /**
  * 根据行、列id集合查询关系映射表
  * @param excelId
  * @param sheetId
  * @param rowList
  * @param colList
  * @return
  */
 List<MRowColCell> getMRowColCellList(String excelId,String sheetId,List<String> rowList,List<String> colList );
 
 
 /**
  * 根据行别名删除关系表
  * @param excelId 数据集
  * @param sheetId 
  * @param field   字段名 row或col
  * @param alais   行或列的别名
  */
 void delMRowColCell(String excelId,String sheetId,String field,String alias);
 
 /**
  * 根据列别名删除关系表
  * @param col     列别名
  * @param excelId 数据集
  */
 void delMRowColCell(Integer col,String excelId);
 
 /**
  * 根据行、列别名集合删除对应关系表
  * @param excelId
  * @param sheetId
  * @param rowList
  * @param colList
  */
 void delMRowColCellList(String excelId,String sheetId,List<String> rowList,List<String> colList );
 
 /**
  * 根据单元格id集合，删除对应的关系表
  * @param excelId
  * @param sheetId
  * @param cellIdList
  */
 void delMRowColCellList(String excelId,String sheetId,List<String> cellIdList );
 
 /**
  * 批量根据cellId修改关系表的cellId
  * @param excelId
  * @param sheetId
  * @param originalId
  * @param modifiedId
  */
 void updateMRowColCell(String excelId,String sheetId,String originalId,String modifiedId);
 
 /**
  * 根据cellID 查询cell列表
  * @param excelId
  * @param sheetId
  * @param cellIdList
  * @return
  */
 List<MCell> getMCellList(String excelId,String sheetId,List<String> cellIdList);
 
 /**
  * 根据MCell的id，返回一个MCell
  * @param excelId
  * @param sheetId
  * @param id
  * @return
  */
 MCell getMCell(String excelId,String sheetId,String id);
 
 /**
  * 根据cellId删除对应MExcelCell对象
  * @param excelId
  * @param sheetId
  * @param cellIdList
  */
 void delMCell(String excelId,String sheetId,List<String> cellIdList);
 
 
 /**
  * 更新隐藏属性
  * @param type
  * @param alias
  * @param state
  * @param excelId
  */
 void updateHidden(String type,String alias,boolean state,String excelId);
 
 /**
  * 根据cell id集合，更新MCell对象中，content中内容
  * @param property
  * @param content
  * @param idList
  * @param excelId
  * @param sheetId
  */
 void updateContent(String property,Object content,List<String> idList,String excelId,String sheetId);
 
 /**
  * 根据CellId更新MCell的合并单元格合并数
  * @param property
  * @param content
  * @param idList
  * @param excelId
  * @param sheetId
  */
 void updateSpan(List<String> cellIdList,String excelId,String sheetId);
 
 
 
}
