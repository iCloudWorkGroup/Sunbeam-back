package com.acmr.excel.dao.impl;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.CriteriaDefinition;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MExcelRowDao;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MExcelColumn;
import com.acmr.excel.model.mongo.MExcelRow;
import com.mongodb.WriteResult;

@Repository("mexcelRowDao")
public class MExcelRowDaoImpl implements MExcelRowDao {
	
    @Resource
	private MongoTemplate mongoTemplate;
	public void delExcelRow(String excelId, String alias) {
		
	 mongoTemplate.remove(new Query(Criteria.where("excelRow.code").is(alias)), MExcelRow.class, excelId);
    
	}
	
	public void updateRowHidden(String excelId,String alias, boolean status) {
		
		Query query = new Query();
		query.addCriteria(Criteria.where("excelRow.code").is(alias));
		Update update = new Update();
		update.set("excelRow.rowhidden", status);
		mongoTemplate.updateFirst(query, update,MExcelRow.class, excelId);
		
	}

	@Override
	public void updateRowHeight(String excelId, String alias, Integer height) {
		Query query = new Query();
		query.addCriteria(Criteria.where("excelRow.code").is(alias));
		Update update = new Update();
		update.set("excelRow.height", height);
		mongoTemplate.updateFirst(query, update,MExcelRow.class, excelId);
		
	}

	@Override
	public MExcelRow getMExcelRow(String excelId, String alias) {
		MExcelRow mexcelRow =  mongoTemplate.findOne(new Query(Criteria.where("excelRow.code").is(alias)), MExcelRow.class,excelId);
		return mexcelRow;
	}

	@Override
	public void updateFont(String name, Object property, String alias, String excelId) {
		Query query = new Query();
	    query.addCriteria(Criteria.where("excelRow.code").in(alias));
	    Update update = new Update();
	    update.set("excelRow.cellstyle.font."+name, property);
	    mongoTemplate.updateMulti(query, update, MExcelCell.class,excelId);
		
	}

}
