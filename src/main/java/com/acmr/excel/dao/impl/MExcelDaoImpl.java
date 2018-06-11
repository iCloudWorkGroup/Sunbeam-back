package com.acmr.excel.dao.impl;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MExcelDao;
import com.acmr.excel.model.mongo.MExcel;

@Repository("mexcelDao")
public class MExcelDaoImpl implements MExcelDao {
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public void updateStep(String excelId, Integer step) {
		
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(excelId));
		Update update = new Update();
		update.set("step", step);
		mongoTemplate.updateFirst(query, update,MExcel.class, excelId);

	}

	@Override
	public void updateFrozen(MExcel mexcel, String excelId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(excelId));
		Update update = new Update();
		update.set("step", mexcel.getStep());
		update.set("viewRowAlias", mexcel.getViewRowAlias());
		update.set("viewColAlias", mexcel.getViewColAlias());
		update.set("rowAlias", mexcel.getRowAlias());
		update.set("colAlias", mexcel.getColAlias());
		update.set("freeze", mexcel.isFreeze());
		mongoTemplate.updateFirst(query, update,MExcel.class, excelId);
		
	}

	@Override
	public void updateUnFrozen(MExcel mexcel, String excelId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(excelId));
		Update update = new Update();
		update.set("step", mexcel.getStep());
		update.set("freeze", mexcel.isFreeze());
		mongoTemplate.updateFirst(query, update,MExcel.class, excelId);
		
	}

	@Override
	public void updateMaxRow(Integer alias, String excelId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(excelId));
		Update update = new Update();
		update.set("maxrow", alias);
		
		mongoTemplate.updateFirst(query, update,MExcel.class, excelId);
		
	}

	@Override
	public void updateMaxCol(Integer alias, String excelId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is(excelId));
		Update update = new Update();
		update.set("maxcol", alias);
		
		mongoTemplate.updateFirst(query, update,MExcel.class, excelId);
		
	}

	public MExcel getMExcel(String excelId){
	     MExcel mExcel = mongoTemplate.findOne(new Query(Criteria.where("_id").is(excelId)), MExcel.class, excelId);//查找sheet属性
	     return mExcel;
	}
}
