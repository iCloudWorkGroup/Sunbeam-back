package com.acmr.excel.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort.Direction;
import org.springframework.data.mongodb.core.IndexOperations;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.index.Index;
import org.springframework.data.mongodb.core.query.BasicQuery;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Service;

import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.RowColList;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MExcelColumn;
import com.acmr.excel.model.mongo.MExcelRow;
import com.acmr.excel.model.mongo.MExcelRowC;
import com.acmr.excel.model.mongo.MExcelSheet;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.util.BinarySearch;
import com.acmr.excel.util.masterworker.Master;
import com.acmr.excel.util.masterworker.Worker;
import com.mongodb.BasicDBObject;
import com.mongodb.DBObject;

import acmr.excel.ExcelHelper;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.excel.pojo.ExcelSheetFreeze;

@Service
public class MongodbServiceImpl {

	@Autowired
	private MongoTemplate mongoTemplate;

	public Object get(Map<String, Object> map, Class<?> clz, String collection) {
		DBObject dbObject = new BasicDBObject();
		dbObject.putAll(map);
		BasicDBObject fieldsObject = new BasicDBObject();
		Query query = new BasicQuery(dbObject, fieldsObject);
		List<?> result = mongoTemplate.find(query, clz, collection);
		return result.size() == 0 ? null : result.get(0);
	}

	public boolean set(String id, Object object) {
		mongoTemplate.insert(object, id);
		return true;
	}

	public boolean update(String excelId, Object object, String express) {
		Update update = new Update();
		update.set(express, object);
		Query query = Query.query(new Criteria().andOperator(Criteria.where(
				"_id").is(excelId)));
		mongoTemplate.upsert(query, update, ExcelBook.class, excelId);
		return true;
	}

	public boolean updateCell(String excelId, Object object) {
		mongoTemplate.save(object, excelId);
		return true;
	}

	public int getStep(String id, Object object) {
		DBObject dbObject = new BasicDBObject();
		dbObject.put("_id", id);
		BasicDBObject fieldsObject = new BasicDBObject();
		fieldsObject.put("step", true);
		Query query = new BasicQuery(dbObject, fieldsObject);
		List<MSheet> result = mongoTemplate.find(query, MSheet.class, id);
		return result.get(0).getStep();
	}
	
	
	public MSheet getMExcel(String excelId){
		MSheet mExcel = mongoTemplate.findOne(new Query(Criteria.where("_id").is(excelId)), MSheet.class, excelId);//查找sheet属性
	     return mExcel;
	}
	
	public ExcelSheet getSheetBySort(int rowBeginSort, int rowEndSort,int colBegin,int colEnd ,String excelId,List<RowCol> sortRclist,List<RowCol> sortClList) {
		
		ExcelSheet excelSheet = new ExcelSheet();
	
		//需要查找的行列
		List<String> rList = new ArrayList<String>();
		List<String> cList = new ArrayList<String>();
		for(int i = rowBeginSort;i<rowEndSort+1;i++){
			rList.add(sortRclist.get(i).getAlias());
		}
		for(int j = colBegin;j<colEnd+1;j++){
			cList.add(sortClList.get(j).getAlias());
		}
	
		//查找行样式
		List<MExcelRow> mRowList = mongoTemplate.find(new Query(Criteria.where("excelRow.code").in(rList)), MExcelRow.class, excelId);
		Map<String, ExcelRow> rowMap = new HashMap<String, ExcelRow>();
		for(MExcelRow mr:mRowList){
		  ExcelRow er =	mr.getExcelRow();
		  rowMap.put(er.getCode(), er);
		}
        //查找列样式
		List<MExcelColumn> mColList = mongoTemplate.find(new Query(Criteria.where("excelColumn.code").in(cList)), MExcelColumn.class, excelId);
		
		Map<String, ExcelColumn> colMap = new HashMap<String, ExcelColumn>();
		for (MExcelColumn mc : mColList) {
			ExcelColumn col = mc.getExcelColumn();
			//excelSheet.getCols().add(col);
			colMap.put(col.getCode(), col);
		}
		//将col对象放入excelSheet
		for(String code:cList){
			ExcelColumn col = colMap.get(code);
			excelSheet.getCols().add(col);
		}
		List<Integer> irList = new ArrayList<Integer>();
		List<Integer> icList = new ArrayList<Integer>();
		stringToInteger(rList,irList);
		stringToInteger(cList,icList);
		//查找关系映射表
		Criteria criatira = new Criteria();
		
		criatira.andOperator(Criteria.where("row").in(irList),
				Criteria.where("col").in(icList));
		List<RowColCell> list = mongoTemplate.find(new Query(criatira), RowColCell.class,excelId);
		//查找对应的cell
		List<String>  inlist = new ArrayList<String>();
		Map<String,String> relationMap = new HashMap<String,String>();//映射关系map
		for(RowColCell rcc:list){
			inlist.add(rcc.getCellId());
			relationMap.put(rcc.getRow()+"_"+rcc.getCol(), rcc.getCellId());
		}
		//查找所有的单元格
		List<MExcelCell> cellList = mongoTemplate.find(new Query(Criteria.where("_id").in(inlist)), MExcelCell.class,excelId);
		Map<String,ExcelCell> cellMap = new HashMap<String,ExcelCell>();
		for(MExcelCell mec:cellList){
			cellMap.put(mec.getId(), mec.getExcelCell());
		}
	
		
		for (int i = 0; i < rList.size(); i++) {
			ExcelRow row = rowMap.get(rList.get(i));
			excelSheet.getRows().add(row);
			List<ExcelCell> cells = new ArrayList<>();
			for (String cl:cList) {
				String cellId = relationMap.get(row.getCode()+"_"+cl);
				cells.add(cellMap.get(cellId));
			}
			row.setCells(cells);
		}

		
		return excelSheet;
	}
	
	public int getRowEndIndex(String excelId,int length,List<RowCol> sortRcList) {
		sortRcList.clear();
		RowColList rowColList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("rList")), RowColList.class, excelId);
		//long ceb2 = System.currentTimeMillis();
		//System.out.println("获得rcList的时间为:" + (ceb2-ceb1));
		if(null == rowColList){
			return 0;
		}
		List<RowCol> rcList = rowColList.getRcList();//得到行列表
		Map<String,RowCol> map = new HashMap<String,RowCol>();
		RowCol rowCol = null;
		
		for(int i =0;i<rcList.size();i++){
			RowCol rc = rcList.get(i);
			if("".equals(rc.getPreAlias())||(null==rc.getPreAlias())){
				rowCol=rc;
				continue;
			}
			map.put(rc.getPreAlias(), rc);
		}
		
		//重新整理行，将行安装展示先后顺序重新排列
		sortRcList.add(rowCol);
		boolean flag = true;
		while(flag){
			rowCol = map.get(rowCol.getAlias());
			if(null == rowCol){
				break;
			}
			sortRcList.add(rowCol);
			if(sortRcList.size()==rcList.size()){//跳出循环
				break;
			}
			 
		}
		
		for(int i=0;i < sortRcList.size();i++) {
			sortRcList.get(i).setTop(getTop(sortRcList, i));
		}
		
		RowCol rowColTop = sortRcList.get(rcList.size() - 1);
		int maxTop = rowColTop.getTop() + rowColTop.getLength();
		
		int endPixel = 0;
		if(length==0){
			endPixel = 0;
		}else if (maxTop < length) {
			endPixel = maxTop;
		} else {
			endPixel =  length + 200;
		}
		int end = BinarySearch.binarySearch(sortRcList, endPixel);
		return end;
	}
	
	
/*	public int getIndexByPixel(String excelId,int pixel,String type) {
		RowColList rowColList = mongoTemplate.findOne(new Query(Criteria.where("_id").is(type)), RowColList.class, excelId);
		List<RowCol> rcList = rowColList.getRcList();
		for(int i=0;i < rcList.size();i++) {
			rcList.get(i).setTop(getTop(rcList, i));
		}
		return BinarySearch.binarySearch(rowColList.getRcList(), pixel);
	}*/
	
	public List<RowCol> getRCList(String excelId,String type){
		RowColList rowColList = mongoTemplate.findOne(new Query(Criteria.where("_id").is(type)), RowColList.class, excelId);
		List<RowCol> rcList = rowColList.getRcList();
		for(int i=0;i < rcList.size();i++) {
			rcList.get(i).setTop(getTop(rcList, i));
		}
		return rcList;
	}
	public int getIndex(List<RowCol> rowColList,int pixel) {
		return BinarySearch.binarySearch(rowColList, pixel);
	}
	
	
	public int getColEndIndex(String excelId,int length,List<RowCol> sortClList) {
		sortClList.clear();
		RowColList colList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("cList")), RowColList.class, excelId);
		//long ceb2 = System.currentTimeMillis();
		//System.out.println("获得rcList的时间为:" + (ceb2-ceb1));
		List<RowCol> cList = colList.getRcList();
		Map<String,RowCol> map = new HashMap<String,RowCol>();
		RowCol rowCol = null;
		
		for(int i =0;i<cList.size();i++){
			RowCol cl = cList.get(i);
			
			if("".equals(cl.getPreAlias())||(null==cl.getPreAlias())){
				rowCol=cl;
				continue;
			}
			map.put(cl.getPreAlias(), cl);
		}
		
		//重新整理行，将行安装展示先后顺序重新排列
		sortClList.add(rowCol);
		boolean flag = true;
		while(flag){
			rowCol = map.get(rowCol.getAlias());
			if(null == rowCol){
				break;
			}
			sortClList.add(rowCol);
			if(sortClList.size()==cList.size()){//跳出循环
				break;
			}
			 
		}
		
		for(int i=0;i < sortClList.size();i++) {
			sortClList.get(i).setTop(getTop(sortClList, i));
		}
		
		RowCol colTop = sortClList.get(sortClList.size() - 1);
		int maxTop = colTop.getTop() + colTop.getLength();
		int endPixel = 0;
		if(length == 0){
			endPixel = 0;
		}else if(maxTop < length) {
			endPixel = maxTop;
		} else {
			endPixel = length + 200;
		}
		int end = BinarySearch.binarySearch(sortClList, endPixel);
		return end;
	}
	
	private int getTop(List<RowCol> rowColList, int i) {
		if (i == 0) {
			return 0;
		}
		RowCol rowCol = rowColList.get(i - 1);
		int tempHeight = rowCol.getLength();
		if(tempHeight == 0){
			tempHeight = -1;
		}

		return rowCol.getTop() + tempHeight + 1;
	}
	
	public boolean saveExcelBook(ExcelBook excelBook, String excelId) {
		if (excelBook == null) {
			return false;
		}
		List<ExcelSheet> excelSheets = excelBook.getSheets();
		ExcelSheet excelSheet = excelSheets.get(0);
		MSheet mExcel = new MSheet();
		mExcel.setId(excelId);
		mExcel.setStep(0);
		mExcel.setMaxrow(excelSheet.getMaxrow());
		mExcel.setMaxcol(excelSheet.getMaxcol());
		mExcel.setSheetName(excelSheet.getName());
		ExcelSheetFreeze freeze = excelSheet.getFreeze();
		if(null != freeze){
		mExcel.setFreeze(freeze.isFreeze());
		mExcel.setViewRowAlias(Integer.toString(freeze.getFirstrow()+1-freeze.getRow()));
		mExcel.setViewColAlias(Integer.toString(freeze.getFirstcol()+1-freeze.getCol()));
		mExcel.setRowAlias(Integer.toString(freeze.getFirstrow()+1));
		mExcel.setColAlias(Integer.toString(freeze.getFirstcol()+1));
		}else{
			mExcel.setFreeze(false);
		}
		set(excelId, mExcel);//存上传文件的最大行和最大列
		List<ExcelColumn> cols = excelSheet.getCols();//得到所有的列
		List<MExcelColumn> tempCols = new ArrayList<MExcelColumn>();
	
		for (int i = 0; i < cols.size(); i++) {
			MExcelColumn mc = new MExcelColumn();
			mc.setExcelColumn(cols.get(i));
			mc.setColSort(i);
			tempCols.add(mc);
		}
		
		IndexOperations iro = mongoTemplate.indexOps(excelId);
		Index rIndex = new Index();
		rIndex.on("rowSort", Direction.ASC);
		iro.ensureIndex(rIndex);
		IndexOperations cellIO = mongoTemplate.indexOps(excelId);
		Index cellIndex = new Index();
		cellIndex.on("rowId", Direction.ASC);
		cellIO.ensureIndex(cellIndex);
		IndexOperations io = mongoTemplate.indexOps(excelId);
		Index index = new Index();
		index.on("colSort", Direction.ASC);
		io.ensureIndex(index);
		
		mongoTemplate.insert(tempCols, excelId);//存列表样式
	

		tempCols = null;
		List<ExcelRow> rows = excelSheet.getRows();
		RowColList rList = new RowColList();
		rList.setId("rList");
		for (int i = 0; i < rows.size(); i++) {
			
			RowCol rc = new RowCol();
			ExcelRow row = rows.get(i);
			rc.setAlias(row.getCode());
			if(row.isRowhidden()){
				rc.setLength(0);
			}else{
				rc.setLength(row.getHeight());
			}
			
			if(i>0){
				
			 rc.setPreAlias(rows.get(i-1).getCode());//存储它前一个值得别名
			}
			rList.getRcList().add(rc);
		}
		mongoTemplate.insert(rList, excelId);//向数据库存贮行信息
		RowColList cList = new RowColList();
		cList.setId("cList");
		for (int i = 0; i < cols.size(); i++) {
			RowCol rc = new RowCol();
			
			ExcelColumn col = cols.get(i);
			rc.setAlias(col.getCode());
			if(col.isColumnhidden()){
				rc.setLength(0);
			}else{
				rc.setLength(col.getWidth());
			}
			
			if(i>0){
				
				rc.setPreAlias(cols.get(i-1).getCode());//存贮它前一行的别名
			}
			cList.getRcList().add(rc);
		}
		mongoTemplate.insert(cList, excelId);//向数据库存贮列信息
		
		/**存贮没有合并单元格的cell对象及行列关系*/
		List<RowColCell> relation = new ArrayList<RowColCell>();
		List<Object> tempList = new ArrayList<Object>();//存cell属性
		
		Master master = new Master(new Worker(), Runtime.getRuntime().availableProcessors(), mongoTemplate, excelId);
		/**找出没有合并的单元格*/
		for (int i = 0; i < rows.size(); i++) {
			
			ExcelRow row = rows.get(i);
			MExcelRow mr = new MExcelRow();
			mr.setExcelRow(row);
			mr.setRowSort(i);
			tempList.add(mr);//存行样式
			List<ExcelCell> cells = row.getCells();
			for (int j = 0; j < cells.size(); j++) {
				
			  if(cells.get(j)!=null && cells.get(j).getColspan()<2 && cells.get(j).getRowspan()<2){//判断是否合并单元格
				MExcelCell mcell = new MExcelCell();
				int rid = i + 1;
				int cid = j + 1;
				mcell.setId(rid + "_" + cid);
				mcell.setRowId(rid + "");
				mcell.setColId(cid + "");
				mcell.setExcelCell(cells.get(j));
				mcell.setRowspan(1);
				mcell.setColspan(1);
				tempList.add(mcell);
				//关系映射表
				RowColCell rcl = new RowColCell();
				rcl.setRow(rid);
				rcl.setCol(cid);
				rcl.setCellId(rid + "_" + cid);
				relation.add(rcl);
                }
			}
			cells.clear();
		}
		List<Sheet> sheets = excelBook.getNativeSheet();//找出合并的单元格
		getMergeCell(sheets,tempList,relation);
		master.submit(tempList);
		master.execute();
		long start = System.currentTimeMillis();
		while(true){
			if(master.isComplete()){
				break;
			}
		}
		long end = System.currentTimeMillis() - start;
		System.out.println("存储时间为:" + end);
		//存贮关系映射表
		mongoTemplate.insert(relation, excelId);
		
		return true;
	}
	

	
	
	
	/***
	 * 获取合并单元格ExcelCell对象及其映射关系RowColCell
	 * @param sheets
	 * @param tempList
	 * @param relation
	 */
	private void getMergeCell(List<Sheet> sheets,List<Object> tempList,List<RowColCell> relation){
		
		for(int i=0;i<sheets.size();i++){
			Sheet sh = sheets.get(i);
			int mergeCount = sh.getNumMergedRegions();
			for(int j=0;j<mergeCount;j++){
				
				CellRangeAddress ca = sh.getMergedRegion(j);
				int firstRow = ca.getFirstRow();
				int firstCol = ca.getFirstColumn();
				int lastRow = ca.getLastRow();
				int lastCol = ca.getLastColumn();
				Row row = sh.getRow(firstRow);
				XSSFCell cell =(XSSFCell) row.getCell(firstCol);
				ExcelSheet es = new ExcelSheet();
				ExcelCell excell = es.getExcelCell(cell);
				excell.setColspan(lastCol-firstCol+1);
				excell.setRowspan(lastRow-firstRow+1);
				
				//合并单元格的最后一个cell，将右、下边框的属性赋值给合并单元格
				Row lrow = sh.getRow(lastRow);
				XSSFCell lastcell =(XSSFCell) lrow.getCell(lastCol);
				XSSFCellStyle cs = lastcell.getCellStyle();
				excell.getCellstyle().setRightborder(ExcelHelper.getExcelBorder(cs.getBorderRight(), cs.getRightBorderXSSFColor()));
		        excell.getCellstyle().setBottomborder(ExcelHelper.getExcelBorder(cs.getBorderBottom(), cs.getBottomBorderXSSFColor()));		
				
				int rid = firstRow + 1;
				int cid = firstCol + 1;
				MExcelCell mcell = new MExcelCell();
				mcell.setId(rid + "_" + cid);
				mcell.setRowId(rid + "");
				mcell.setColId(cid + "");
				mcell.setExcelCell(excell);
				mcell.setRowspan(lastRow-firstRow+1);
				mcell.setColspan(lastCol-firstCol+1);
				tempList.add(mcell);
				//关系映射表
				for(int k=firstRow;k<lastRow+1;k++){
					for(int l=firstCol;l<lastCol+1;l++){
						//多个单元格指向一个单元对象
						RowColCell rcl = new RowColCell();
						rcl.setRow(k+1);
						rcl.setCol(l+1);
						rcl.setCellId(rid + "_" + cid);
						relation.add(rcl);
					}
				}
				
			}
		}
		
	}
	
	/***
	 * 将字符串类型的List转换成Integer类型的List
	 * @param oldl
	 * @param newl
	 */
	private void stringToInteger(List<String> oldl,List<Integer> newl){
		
		for(String old:oldl){
			
			newl.add(Integer.valueOf(old));
		}
	}
	
	/***
	 * 得到经过排序之后的列
	 * @param sortClList
	 * @param newClMap
	 * @param excelId
	 */
	public void getColList(List<RowCol> sortClList,Map<String,RowCol> newClMap,String excelId){
		
		RowColList colList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("cList")), RowColList.class, excelId);
		List<RowCol> cList = colList.getRcList();
		Map<String,RowCol> map = new HashMap<String,RowCol>();
		RowCol rowCol = null;
		
		for(int i =0;i<cList.size();i++){
			RowCol cl = cList.get(i);
			
			if("".equals(cl.getPreAlias())||(null==cl.getPreAlias())){
				rowCol=cl;
				continue;
			}
			map.put(cl.getPreAlias(), cl);
		}
		
		//重新整理行，将行安装展示先后顺序重新排列
		sortClList.add(rowCol);
		boolean flag = true;
		while(flag){
			rowCol = map.get(rowCol.getAlias());
			sortClList.add(rowCol);
			if(sortClList.size()==cList.size()){//跳出循环
				break;
			}
			 
		}
		
		for(int i=0;i < sortClList.size();i++) {
			sortClList.get(i).setTop(getTop(sortClList, i));
			newClMap.put(sortClList.get(i).getAlias(), sortClList.get(i));
		}
		
	}
	
    /***
     * 得到排序之后的行 
     * @param sortRclist
     * @param newRcMap
     * @param excelId
     */
   public void getRowList(List<RowCol> sortRclist,Map<String,RowCol> newRcMap,String excelId){
		
		RowColList rowColList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("rList")), RowColList.class, excelId);
		//long ceb2 = System.currentTimeMillis();
		//System.out.println("获得rcList的时间为:" + (ceb2-ceb1));
		List<RowCol> rcList = rowColList.getRcList();//得到行列表
		Map<String,RowCol> map = new HashMap<String,RowCol>();
		RowCol rowCol = null;
		
		for(int i =0;i<rcList.size();i++){
			RowCol rc = rcList.get(i);
			if("".equals(rc.getPreAlias())||(null==rc.getPreAlias())){
				rowCol=rc;
				continue;
			}
			map.put(rc.getPreAlias(), rc);
		}
		
		//重新整理行，将行安装展示先后顺序重新排列
		sortRclist.add(rowCol);
		boolean flag = true;
		while(flag){
			rowCol = map.get(rowCol.getAlias());
			sortRclist.add(rowCol);
			if(sortRclist.size()==rcList.size()){//跳出循环
				break;
			}
			 
		}
		
		for(int i=0;i < sortRclist.size();i++) {
			sortRclist.get(i).setTop(getTop(sortRclist, i));
			newRcMap.put(sortRclist.get(i).getAlias(), sortRclist.get(i));
		}
		
	}
	
	public int getIndexByPixel(List<RowCol> sortRclist,int pixel){
		
		RowCol rowColTop = sortRclist.get(sortRclist.size() - 1);
		int maxTop = rowColTop.getTop() + rowColTop.getLength();
		
		int endPixel = 0;
		if(pixel==0){
			endPixel = 0;
		}else if (maxTop < pixel) {
			endPixel = maxTop;
		} else {
			endPixel =  pixel;
		}
		int end = BinarySearch.binarySearch(sortRclist, endPixel);
		return end;
	}
	/**
	 * 根据排序之后的行列指针，查询列、行、单元格等对象
	 * @param rowBeginSort
	 * @param rowEndSort
	 * @param colBegin
	 * @param colEnd
	 * @param excelId
	 * @param sortRclist
	 * @param sortClList
	 * @return
	 */
	public MExcelSheet getSheetByIndex(int rowBeginIndex, int rowEndIndex,
			int colBeginIndex,int colEndIndex ,String excelId,List<RowCol> sortRclist,
			List<RowCol> sortClList) {
		MSheet mExcel = mongoTemplate.findOne(new Query(Criteria.where("_id").is(excelId)), 
				MSheet.class, excelId);//查找sheet属性
		MExcelSheet excelSheet = new MExcelSheet();
		excelSheet.setMaxcol(mExcel.getMaxcol());
		excelSheet.setMaxrow(mExcel.getMaxrow());
		excelSheet.setName(mExcel.getSheetName());
		List<String> rList = new ArrayList<String>();
		List<String> cList = new ArrayList<String>();
		for(int i = rowBeginIndex;i<rowEndIndex+1;i++){
			rList.add(sortRclist.get(i).getAlias());
		}
		for(int j = colBeginIndex;j<colEndIndex+1;j++){
			cList.add(sortClList.get(j).getAlias());
		}
	
		//查找行样式
		List<MExcelRow> mRowList = mongoTemplate.find(new Query(Criteria.where("excelRow.code").in(rList)), MExcelRow.class, excelId);
        //查找列样式
		List<MExcelColumn> mColList = mongoTemplate.find(new Query(Criteria.where("excelColumn.code").in(cList)), MExcelColumn.class, excelId);
		
		for (MExcelColumn mc : mColList) {
			ExcelColumn col = mc.getExcelColumn();
			excelSheet.getCols().add(col);
			
		}
		
		List<Integer> irList = new ArrayList<Integer>();
		List<Integer> icList = new ArrayList<Integer>();
		stringToInteger(rList,irList);
		stringToInteger(cList,icList);
		//查找关系映射表
		Criteria criatira = new Criteria();
		
		criatira.andOperator(Criteria.where("row").in(irList),
				Criteria.where("col").in(icList));
		List<RowColCell> list = mongoTemplate.find(new Query(criatira), RowColCell.class,excelId);
		//查找对应的cell
		List<String>  inlist = new ArrayList<String>();
		Map<String,String> relationMap = new HashMap<String,String>();//映射关系map
		for(RowColCell rcc:list){
			inlist.add(rcc.getCellId());
			relationMap.put(rcc.getRow()+"_"+rcc.getCol(), rcc.getCellId());
		}
		//查找所有的单元格
		List<MExcelCell> cellList = mongoTemplate.find(new Query(Criteria.where("_id").in(inlist)), MExcelCell.class,excelId);
		Map<String,MExcelCell> cellMap = new HashMap<String,MExcelCell>();
		for(MExcelCell mec:cellList){
			cellMap.put(mec.getId(), mec);
		}
	
		
		for (int i = 0; i < mRowList.size(); i++) {
			ExcelRow row = mRowList.get(i).getExcelRow();
			MExcelRowC rowC = new MExcelRowC(row);
			excelSheet.getRows().add(rowC);
			List<MExcelCell> cells = new ArrayList<MExcelCell>();
			for (String cl:cList) {
				String cellId = relationMap.get(row.getCode()+"_"+cl);
				cells.add(cellMap.get(cellId));
			}
			rowC.setCells(cells);
		}

		
		return excelSheet;
	}
	
}
