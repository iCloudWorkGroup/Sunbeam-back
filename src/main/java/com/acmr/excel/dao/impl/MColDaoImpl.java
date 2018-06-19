package com.acmr.excel.dao.impl;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MColDao;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MExcelColumn;
import com.acmr.excel.model.mongo.MRow;

@Repository("mcolDao")
public class MColDaoImpl implements MColDao {
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<MCol> getMColList(String excelId, String sheetId, List<String> colList) {
		List<MCol> mColList = mongoTemplate.find(new Query(Criteria.where("alias").in(colList).and("sheetId").is(sheetId).and("_class").is(MCol.class.getName())), MCol.class, excelId);
		return mColList;
	}
	
	@Override
	public List<MCol> getMColList(String excelId, String sheetId) {
		List<MCol> mColList = mongoTemplate.find(new Query(Criteria.where("sheetId").is(sheetId).and("_class").is(MCol.class.getName())), MCol.class, excelId);
		return mColList;
	}

	@Override
	public void delExcelCol(String excelId, String sheetId, String alias) {
		
		mongoTemplate.remove(new Query(Criteria.where("alias").is(alias).and("_class").is(MCol.class.getName()).and("sheetId").is(sheetId)),
				MCol.class, excelId);
		
	}

	
	public void updateColHiddenStatus(String excelId, String sheetId, String alias, boolean status) {
		Query query = new Query();
		query.addCriteria(Criteria.where("alias").is(alias).and("_class").is(MCol.class.getName()).and("sheetId").is(sheetId));
		Update update = new Update();
		update.set("hidden", status);
		mongoTemplate.updateFirst(query, update,MCol.class, excelId);
		
	}

	@Override
	public MCol getMCol(String excelId, String sheetId, String alias) {
		MCol mcol =  mongoTemplate.findOne(new Query(Criteria.where("alias").is(alias).and("_class").is(MCol.class.getName()).and("sheetId").is(sheetId)),
				MCol.class,excelId);
		return mcol;
		
	}

	@Override
	public void updateColWidth(String excelId, String sheetId, String alias, Integer width) {
		
		Query query = new Query();
		query.addCriteria(Criteria.where("alias").is(alias).and("_class").is(MCol.class.getName()).and("sheetId").is(sheetId));
		Update update = new Update();
		update.set("width", width);
		mongoTemplate.updateFirst(query, update,MCol.class, excelId);
		
	}

	@Override
	public void updateContent(String property, Object content, List<String> aliasList, String excelId, String sheetId) {
		Query query = new Query();
	    query.addCriteria(Criteria.where("alias").in(aliasList).and("_class").is(MCol.class.getName()).and("sheetId").is(sheetId));
	    Update update = new Update();
	    update.set("props.content."+property, content);
	    mongoTemplate.updateMulti(query, update, MCol.class,excelId);
		
	}

}
