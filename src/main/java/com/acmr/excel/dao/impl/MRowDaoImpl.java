package com.acmr.excel.dao.impl;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MRowDao;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MExcelRow;
import com.acmr.excel.model.mongo.MRow;

@Repository("mrowDao")
public class MRowDaoImpl implements MRowDao {
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<MRow> getMRowList(String excelId, String sheetId, List<String> rowList) {
		List<MRow> mRowList = mongoTemplate.find(new Query(Criteria.where("alias").in(rowList).and("sheetId").is(sheetId).and("_class").is(MRow.class.getName())), MRow.class, excelId);
		return mRowList;
	}
	
	@Override
	public List<MRow> getMRowList(String excelId, String sheetId) {
		List<MRow> mRowList = mongoTemplate.find(new Query(Criteria.where("sheetId").is(sheetId).and("_class").is(MRow.class.getName())), MRow.class, excelId);
		return mRowList;
	}

	@Override
	public void delMRow(String excelId, String sheetId, String alias) {
		
		mongoTemplate.remove(new Query(Criteria.where("alias").is(alias).and("_class").is(MRow.class.getName()).and("sheetId").is(sheetId)), 
				MRow.class, excelId);
		
	}

	@Override
	public void updateRowHidden(String excelId, String sheetId, String alias, boolean status) {
		Query query = new Query();
		query.addCriteria(Criteria.where("alias").is(alias).and("_class").is(MRow.class.getName()).and("sheetId").is(sheetId));
		Update update = new Update();
		update.set("hidden", status);
		mongoTemplate.updateFirst(query, update,MRow.class, excelId);
		
	}

	@Override
	public MRow getMRow(String excelId, String sheetId, String alias) {
		MRow mrow =  mongoTemplate.findOne(new Query(Criteria.where("alias").is(alias).and("_class").is(MRow.class.getName()).and("sheetId").is(sheetId)),
				MRow.class,excelId);
		return mrow;
	}

	@Override
	public void updateRowHeight(String excelId, String sheetId, String alias, Integer height) {
		Query query = new Query();
		query.addCriteria(Criteria.where("alias").is(alias).and("_class").is(MRow.class.getName()).and("sheetId").is(sheetId));
		Update update = new Update();
		update.set("height", height);
		mongoTemplate.updateFirst(query, update,MRow.class, excelId);
		
		
	}

	@Override
	public void updateContent(String property, Object content, List<String> aliasList, String excelId, String sheetId) {
		
		Query query = new Query();
	    query.addCriteria(Criteria.where("alias").in(aliasList).and("_class").is(MRow.class.getName()).and("sheetId").is(sheetId));
	    Update update = new Update();
	    update.set("props.content."+property, content);
	    mongoTemplate.updateMulti(query, update, MRow.class,excelId);
		
	}


}
