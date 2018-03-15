package com.acmr.excel.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort.Direction;
import org.springframework.data.mongodb.core.IndexOperations;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.index.Index;
import org.springframework.data.mongodb.core.query.BasicQuery;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Meta;
import org.springframework.data.mongodb.core.query.Order;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Service;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;

import com.acmr.excel.model.complete.Gly;
import com.acmr.excel.model.complete.ReturnParam;
import com.acmr.excel.model.mongo.MExcel;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MExcelColumn;
import com.acmr.excel.model.mongo.MExcelRow;
import com.acmr.excel.model.position.RowCol;
import com.acmr.excel.model.position.RowColList;
import com.acmr.excel.service.StoreService;
import com.acmr.excel.util.BinarySearch;
import com.acmr.excel.util.masterworker.Master;
import com.acmr.excel.util.masterworker.Worker;
import com.mongodb.BasicDBObject;
import com.mongodb.DBObject;

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
		List<MExcel> result = mongoTemplate.find(query, MExcel.class, id);
		return result.get(0).getStep();
	}
	
	
	public void getParam(String excelId,ReturnParam returnParam) {
		RowColList rowColList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("rList")), RowColList.class, excelId);
		List<RowCol> rcList = rowColList.getRcList();
		for(int i=0;i < rcList.size();i++) {
			rcList.get(i).setTop(getTop(rcList, i));
		}
		RowColList colList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("cList")), RowColList.class, excelId);
		List<RowCol> cList = colList.getRcList();
		for(int i=0;i < cList.size();i++) {
			cList.get(i).setTop(getTop(cList, i));
		}
		returnParam.setMaxPixel(rcList.get(rcList.size()-1).getTop());
		returnParam.setMaxPixel(rcList.get(rcList.size()-1).getTop());
		returnParam.setMaxColPixel(cList.get(cList.size()-1).getTop());
	}
	
	
	
	
//	public ExcelSheet getSheetBySort(int rowBeginSort, int rowEndSort,int colBegin,int colEnd ,String excelId) {
//		MExcel mExcel = mongoTemplate.findOne(new Query(Criteria.where("_id").is(excelId)), MExcel.class, excelId);
//		long rb1 = System.currentTimeMillis();
//		List<MExcelRow> mRowList = mongoTemplate.find(new Query(Criteria.where("rowSort").gte(rowBeginSort)
//						.lte(rowEndSort)), MExcelRow.class, excelId);
//		long rb2 = System.currentTimeMillis();
//		long cb1 = System.currentTimeMillis();
//		List<MExcelColumn> mColList = mongoTemplate.find(new Query(Criteria.where("colSort").gte(colBegin).lte(colEnd)), 
//				MExcelColumn.class, excelId);
//		long cb2 = System.currentTimeMillis();
//		ExcelSheet excelSheet = new ExcelSheet();
//		Map<String, MExcelColumn> colMap = new HashMap<String, MExcelColumn>();
//		for (MExcelColumn mc : mColList) {
//			ExcelColumn col = mc.getExcelColumn();
//			excelSheet.getCols().add(col);
//			colMap.put(col.getCode(), mc);
//		}
//		excelSheet.setMaxcol(mExcel.getMaxcol());
//		excelSheet.setMaxrow(mExcel.getMaxrow());
//		excelSheet.setName(mExcel.getSheetName());
//		long cell1 = System.currentTimeMillis();
//		
//		for (int i = 0; i < mRowList.size(); i++) {
//			ExcelRow row = mRowList.get(i).getExcelRow();
//			excelSheet.getRows().add(row);
//			long ceb1 = System.currentTimeMillis();
////			List<MExcelCell> cellList = mongoTemplate.find(new Query(Criteria.where("rowId").is(row.getCode())), MExcelCell.class,
////					excelId);
//			List<MExcelCell> cellList = mongoTemplate.find(new Query(Criteria.where("crSort").gte(rowBeginSort)
//					.lte(rowEndSort)), MExcelCell.class,
//					excelId);
//			long ceb2 = System.currentTimeMillis();
//			System.out.println("查询cell=========================="+(ceb2 - ceb1));
//			List<ExcelCell> cells = new ArrayList<>();
//			row.setCells(cells);
//			long cc1 = System.currentTimeMillis();
//			for (MExcelCell mc : cellList) {
//				MExcelColumn mec = colMap.get(mc.getColId());
//				//System.out.println(mc.getColId() + "=====" + mec);
//				if (mec == null) {
//					cells.add(null);
//				}else {
//					cells.add(mc.getExcelCell());
//				}
//			}
//			long cc2 = System.currentTimeMillis();
//			System.out.println("组装cell=========================="+(cc2 - cc1));
//		}
//		long cell2 = System.currentTimeMillis();
//		System.out.println("row=========================="+(rb2 - rb1));
//		System.out.println("col=========================="+(cb2 - cb1));
//		System.out.println("cell=========================="+(cell2 - cell1));
//		System.out.println("==================================================");
//		return excelSheet;
//	}
	public ExcelSheet getSheetBySort(int rowBeginSort, int rowEndSort,int colBegin,int colEnd ,String excelId) {
		MExcel mExcel = mongoTemplate.findOne(new Query(Criteria.where("_id").is(excelId)), MExcel.class, excelId);
//		long rb1 = System.currentTimeMillis();
		List<MExcelRow> mRowList = mongoTemplate.find(new Query(Criteria.where("rowSort").gte(rowBeginSort)
						.lte(rowEndSort)), MExcelRow.class, excelId);
//		long rb2 = System.currentTimeMillis();
//		long cb1 = System.currentTimeMillis();
		List<MExcelColumn> mColList = mongoTemplate.find(new Query(Criteria.where("colSort").gte(colBegin).lte(colEnd)), 
				MExcelColumn.class, excelId);
//		long cb2 = System.currentTimeMillis();
		ExcelSheet excelSheet = new ExcelSheet();
		Map<String, MExcelColumn> colMap = new HashMap<String, MExcelColumn>();
		for (MExcelColumn mc : mColList) {
			ExcelColumn col = mc.getExcelColumn();
			excelSheet.getCols().add(col);
			colMap.put(col.getCode(), mc);
		}
		excelSheet.setMaxcol(mExcel.getMaxcol());
		excelSheet.setMaxrow(mExcel.getMaxrow());
		excelSheet.setName(mExcel.getSheetName());
		
		for (int i = 0; i < mRowList.size(); i++) {
			ExcelRow row = mRowList.get(i).getExcelRow();
			excelSheet.getRows().add(row);
			List<ExcelCell> cells = new ArrayList<>();
			for (int j = 0; j < 100; j++) {
				cells.add(null);
			}
			row.setCells(cells);
		}
//			long ceb1 = System.currentTimeMillis();
//			List<MExcelCell> cellList = mongoTemplate.find(new Query(Criteria.where("rowId").is(row.getCode())), MExcelCell.class,
//					excelId);
			List<MExcelCell> cellList = mongoTemplate.find(new Query(Criteria.where("crSort").gte(rowBeginSort)
					.lte(rowEndSort)), MExcelCell.class,
					excelId);
//			long ceb2 = System.currentTimeMillis();
//			System.out.println("查询cell=========================="+(ceb2 - ceb1));
			
//			long cc1 = System.currentTimeMillis();
			for (MExcelCell mc : cellList) {
				MExcelColumn mec = colMap.get(mc.getColId());
				//System.out.println(mc.getColId() + "=====" + mec);
				if (mec == null) {
					//excelSheet.getRows().get(mc.getCrSort()).add(null);
				}else {
					ExcelRow row = excelSheet.getRows().get(mc.getCrSort()-rowBeginSort);
					row.getCells().set(mc.getCclSort(),mc.getExcelCell());
				}
			}
//			long cc2 = System.currentTimeMillis();
//			System.out.println("组装cell=========================="+(cc2 - cc1));
//		
//		System.out.println("row=========================="+(rb2 - rb1));
//		System.out.println("col=========================="+(cb2 - cb1));
//		System.out.println("==================================================");
		return excelSheet;
	}
	
	public int getRowEndIndex(String excelId,int length) {
		//long ceb1 = System.currentTimeMillis();
		RowColList rowColList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("rList")), RowColList.class, excelId);
		//long ceb2 = System.currentTimeMillis();
		//System.out.println("获得rcList的时间为:" + (ceb2-ceb1));
		List<RowCol> rcList = rowColList.getRcList();
		for(int i=0;i < rcList.size();i++) {
			rcList.get(i).setTop(getTop(rcList, i));
		}
		int minTop = rcList.get(0).getTop();
		RowCol rowColTop = rcList.get(rcList.size() - 1);
		int maxTop = rowColTop.getTop() + rowColTop.getLength();
		int startAlaisPixel = 0;                                 
		int Offset = startAlaisPixel - minTop;
		int startPixel = Offset < 200 ? Offset : startAlaisPixel - 200;
		int endPixel = 0;
		if (maxTop < length) {
			endPixel = maxTop;
		} else {
			endPixel = startPixel + length + 200;
		}
		int end = BinarySearch.binarySearch(rcList, endPixel);
		return end;
	}
	
	
	public int getIndexByPixel(String excelId,int pixel,String type) {
		RowColList rowColList = mongoTemplate.findOne(new Query(Criteria.where("_id").is(type)), RowColList.class, excelId);
		List<RowCol> rcList = rowColList.getRcList();
		for(int i=0;i < rcList.size();i++) {
			rcList.get(i).setTop(getTop(rcList, i));
		}
		return BinarySearch.binarySearch(rowColList.getRcList(), pixel);
	}
	
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
	
	
	public int getColEndIndex(String excelId,int length) {
		//long ceb1 = System.currentTimeMillis();
		RowColList colList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("cList")), RowColList.class, excelId);
		//long ceb2 = System.currentTimeMillis();
		//System.out.println("获得rcList的时间为:" + (ceb2-ceb1));
		List<RowCol> cList = colList.getRcList();
		for(int i=0;i < cList.size();i++) {
			cList.get(i).setTop(getTop(cList, i));
		}
		int minTop = cList.get(0).getTop();
		RowCol colTop = cList.get(cList.size() - 1);
		int maxTop = colTop.getTop() + colTop.getLength();
		int startAlaisPixel = 0;                                 
		int Offset = startAlaisPixel - minTop;
		int startPixel = Offset < 200 ? Offset : startAlaisPixel - 200;
		int endPixel = 0;
		if (maxTop < length) {
			endPixel = maxTop;
		} else {
			endPixel = startPixel + length + 200;
		}
		int end = BinarySearch.binarySearch(cList, endPixel);
		return end;
	}
	private int getTop(List<RowCol> rowColList, int i) {
		if (i == 0) {
			return 0;
		}
		RowCol rowCol = rowColList.get(i - 1);
		int tempHeight = rowCol.getLength();
//		if(gly.isHidden()){
//			tempHeight = -1;
//		}
		return rowCol.getTop() + tempHeight + 1;
	}
	
	
	
	
	
	
	

	public boolean saveExcelBook(ExcelBook excelBook, String excelId) {
		if (excelBook == null) {
			return false;
		}
		List<ExcelSheet> excelSheets = excelBook.getSheets();
		ExcelSheet excelSheet = excelSheets.get(0);
		MExcel mExcel = new MExcel();
		mExcel.setId(excelId);
		mExcel.setStep(0);
		mExcel.setMaxrow(excelSheet.getMaxrow());
		mExcel.setMaxcol(excelSheet.getMaxcol());
		set(excelId, mExcel);
		List<ExcelColumn> cols = excelSheet.getCols();
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
		IndexOperations crio = mongoTemplate.indexOps(excelId);
		Index crindex = new Index();
		crindex.on("crSort", Direction.ASC);
		crio.ensureIndex(crindex);
		IndexOperations ccio = mongoTemplate.indexOps(excelId);
		Index ccindex = new Index();
		ccindex.on("cclSort", Direction.ASC);
		ccio.ensureIndex(ccindex);
		long b1 = System.currentTimeMillis();
		mongoTemplate.insert(tempCols, excelId);
		long b2 = System.currentTimeMillis();
		tempCols = null;
		List<ExcelRow> rows = excelSheet.getRows();
		RowColList rList = new RowColList();
		rList.setId("rList");
		for (int i = 0; i < rows.size(); i++) {
			RowCol rc = new RowCol();
			ExcelRow row = rows.get(i);
			rc.setAlias(row.getCode());
			rc.setLength(row.getHeight());
			rList.getRcList().add(rc);
		}
		mongoTemplate.insert(rList, excelId);
		RowColList cList = new RowColList();
		cList.setId("cList");
		for (int i = 0; i < cols.size(); i++) {
			RowCol rc = new RowCol();
			ExcelColumn col = cols.get(i);
			rc.setAlias(col.getCode());
			rc.setLength(col.getWidth());
			cList.getRcList().add(rc);
		}
		mongoTemplate.insert(cList, excelId);
		Master master = new Master(new Worker(), Runtime.getRuntime().availableProcessors(), mongoTemplate, excelId);
		for (int i = 0; i < rows.size(); i++) {
			List<Object> tempList = new ArrayList<Object>();
			ExcelRow row = rows.get(i);
			MExcelRow mr = new MExcelRow();
			mr.setExcelRow(row);
			mr.setRowSort(i);
			tempList.add(mr);
			List<ExcelCell> cells = row.getCells();
			for (int j = 0; j < cells.size(); j++) {
				MExcelCell mcell = new MExcelCell();
				int rid = i + 1;
				int cid = j + 1;
				mcell.setId(rid + "_" + cid);
				mcell.setRowId(rid + "");
				mcell.setColId(cid + "");
				mcell.setExcelCell(cells.get(j));
				mcell.setCrSort(i);
				mcell.setCclSort(j);
				tempList.add(mcell);
			}
			cells.clear();
			master.submit(tempList);
		}
		master.execute();
		long start = System.currentTimeMillis();
		while(true){
			if(master.isComplete()){
				break;
			}
		}
		long end = System.currentTimeMillis() - start;
		System.out.println("存储时间为:" + end);
		//System.out.println("cpu线程数:  " + Runtime.getRuntime().availableProcessors());
		return true;
	}
	public static void main(String[] args) {
		System.out.println();
	}
}
