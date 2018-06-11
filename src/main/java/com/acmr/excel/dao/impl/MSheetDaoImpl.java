package com.acmr.excel.dao.impl;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.model.mongo.MSheet;

@Repository("msheetDao")
public class MSheetDaoImpl implements MSheetDao {
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public void updateStep(String excelId,String sheetId, Integer step) {
		
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(sheetId));
		Update update = new Update();
		update.set("step", step);
		mongoTemplate.updateFirst(query, update,MSheet.class, excelId);

	}

	@Override
	public void updateFrozen(MSheet msheet, String excelId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(excelId));
		Update update = new Update();
		update.set("step", msheet.getStep());
		update.set("viewRowAlias", msheet.getViewRowAlias());
		update.set("viewColAlias", msheet.getViewColAlias());
		update.set("rowAlias", msheet.getRowAlias());
		update.set("colAlias", msheet.getColAlias());
		update.set("freeze", msheet.getFreeze());
		mongoTemplate.updateFirst(query, update,MSheet.class, excelId);
		
	}

	@Override
	public void updateUnFrozen(MSheet msheet, String excelId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(excelId));
		Update update = new Update();
		update.set("step", msheet.getStep());
		update.set("freeze", msheet.getFreeze());
		mongoTemplate.updateFirst(query, update,MSheet.class, excelId);
		
	}

	@Override
	public void updateMaxRow(Integer alias, String excelId,String sheetId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(sheetId));
		Update update = new Update();
		update.set("maxrow", alias);
		
		mongoTemplate.updateFirst(query, update,MSheet.class, excelId);
		
	}

	@Override
	public void updateMaxCol(Integer alias, String excelId,String sheetId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(sheetId));
		Update update = new Update();
		update.set("maxcol", alias);
		
		mongoTemplate.updateFirst(query, update,MSheet.class, excelId);
		
	}

	public MSheet getMSheet(String excelId,String sheetId){
		MSheet msheet = mongoTemplate.findOne(new Query(Criteria.where("_id").is(sheetId)), MSheet.class, excelId);//查找sheet属性
	     return msheet;
	}
}
