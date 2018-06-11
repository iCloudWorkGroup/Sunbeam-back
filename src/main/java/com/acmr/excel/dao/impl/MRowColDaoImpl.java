package com.acmr.excel.dao.impl;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.RowColList;
import com.acmr.excel.model.mongo.MRowColList;
import com.mongodb.BasicDBObject;

@Repository("mrowColDao")
public class MRowColDaoImpl  implements MRowColDao{
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public void getRowList(List<RowCol> sortRcList, String excelId,String sheetId) {
		
		MRowColList rowColList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("rList").and("sheetId").is(sheetId)), MRowColList.class, excelId);
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
		
	}

	@Override
	public void getColList(List<RowCol> sortClList, String excelId,String sheetId) {
		
		MRowColList colList = mongoTemplate.findOne(new Query(Criteria.where("_id").is("cList").and("sheetId").is(sheetId)), MRowColList.class, excelId);
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
		
	}
	
	
	private int getTop(List<RowCol> rowColList, int i) {
		if (i == 0) {
			return 0;
		}
		RowCol rowCol = rowColList.get(i - 1);
		int tempHeight = rowCol.getLength();

		return rowCol.getTop() + tempHeight + 1;
	}

	@Override
	public void updateRowCol(String excelId,String sheetId,String id,String alias,String preAlias) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(id).and("rcList.alias").is(alias).and("sheetId").is(sheetId));
		Update update = new Update();
		update.set("rcList.$.preAlias", preAlias);
		mongoTemplate.updateFirst(query, update, excelId);
		
	}

	@Override
	public void insertRowCol(String excelId,String sheetId, RowCol rowCol, String id) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(id).and("sheetId").is(sheetId));
		Update update = new Update();
		update.addToSet("rcList",rowCol);
		mongoTemplate.upsert(query, update, excelId);
		
		
	}

	@Override
	public void delRowCol(String excelId,String sheetId, String alias, String id) {
		Query query = Query.query(Criteria.where("_id").is(id).and("sheetId").is(sheetId));
		BasicDBObject s = new BasicDBObject();
		s.put("alias",alias);
		Update update = new Update();
		update.pull("rcList", s);
	    mongoTemplate.updateFirst(query, update, excelId);
		
	}

	@Override
	public void updateRowColLength(String excelId, String id, String alias, Integer length) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(id).and("rcList.alias").is(alias));
		Update update = new Update();
		update.set("rcList.$.length", length);
		mongoTemplate.updateFirst(query, update, excelId);
		
	}

}
