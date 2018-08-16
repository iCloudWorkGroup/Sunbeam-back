package com.acmr.excel.dao.impl;

import java.io.Serializable;
import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MColDao;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MExcelColumn;
import com.acmr.excel.model.mongo.MRow;

@Repository("mcolDao")
public class MColDaoImpl implements MColDao,Serializable {

	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<MCol> getMColList(String excelId, String sheetId,
			List<String> colList) {
		List<MCol> mColList = mongoTemplate.find(
				new Query(Criteria.where("alias").in(colList).and("sheetId")
						.is(sheetId).and("_class").is(MCol.class.getName())),
				MCol.class, excelId);
		return mColList;
	}

	@Override
	public List<MCol> getMColList(String excelId, String sheetId) {
		List<MCol> mColList = mongoTemplate.find(
				new Query(Criteria.where("sheetId").is(sheetId).and("_class")
						.is(MCol.class.getName())),
				MCol.class, excelId);
		return mColList;
	}

	@Override
	public void delMCol(String excelId, String sheetId, String alias) {

		mongoTemplate.remove(
				new Query(Criteria.where("alias").is(alias).and("_class")
						.is(MCol.class.getName()).and("sheetId").is(sheetId)),
				MCol.class, excelId);

	}
	
	@Override
	public void delMColList(String excelId, String sheetId,
			List<String> aliasList) {
		mongoTemplate.remove(
				new Query(Criteria.where("alias").in(aliasList).and("_class")
						.is(MCol.class.getName()).and("sheetId").is(sheetId)),
				MCol.class, excelId);
		
	}

	public void updateColHiddenStatus(String excelId, String sheetId,
			String alias, Boolean status) {
		Query query = new Query();
		query.addCriteria(Criteria.where("alias").is(alias).and("_class")
				.is(MCol.class.getName()).and("sheetId").is(sheetId));
		Update update = new Update();
		update.set("hidden", status);
		mongoTemplate.updateFirst(query, update, MCol.class, excelId);

	}

	@Override
	public MCol getMCol(String excelId, String sheetId, String alias) {
		MCol mcol = mongoTemplate.findOne(
				new Query(Criteria.where("alias").is(alias).and("_class")
						.is(MCol.class.getName()).and("sheetId").is(sheetId)),
				MCol.class, excelId);
		return mcol;

	}

	@Override
	public void updateColWidth(String excelId, String sheetId, String alias,
			Integer width) {

		Query query = new Query();
		query.addCriteria(Criteria.where("alias").is(alias).and("_class")
				.is(MCol.class.getName()).and("sheetId").is(sheetId));
		Update update = new Update();
		update.set("width", width);
		mongoTemplate.updateFirst(query, update, MCol.class, excelId);

	}

	@Override
	public void updateContent(String property1, Object value1, String property2,
			Object value2, List<String> aliasList, String excelId,
			String sheetId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("alias").in(aliasList).and("_class")
				.is(MCol.class.getName()).and("sheetId").is(sheetId));
		Update update = new Update();
		if(null!=property2){
			update.set("props.content." + property1, value1);
			update.set("props.content." + property2, value2);
		}else{
			update.set("props.content." + property1, value1);
		}
		mongoTemplate.updateMulti(query, update, MCol.class, excelId);

	}
	
	@Override
	public void updateBorderList(String property, Object value, List<String> aliasList,
			String excelId, String sheetId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("alias").in(aliasList).and("_class")
				.is(MCol.class.getName()).and("sheetId").is(sheetId));
		Update update = new Update();
		if("all".equals(property)){
			update.set("props.border.left" , value);
			update.set("props.border.top", value);
			update.set("props.border.right" , value);
			update.set("props.border.bottom", value);
		}else if("none".equals(property)){
			update.set("props.border.left" , value);
			update.set("props.border.top", value);
			update.set("props.border.right" , value);
			update.set("props.border.bottom", value);
		}else{
		    update.set("props.border." + property, value);
		}
		mongoTemplate.updateMulti(query, update, MCol.class, excelId);
		
	}

	@Override
	public void updateBorder(String property, Object value, String alias,
			String excelId, String sheetId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("alias").is(alias).and("_class")
				.is(MCol.class.getName()).and("sheetId").is(sheetId));
		Update update = new Update();
		if("all".equals(property)){
			update.set("props.border.left" , value);
			update.set("props.border.top", value);
			update.set("props.border.right" , value);
			update.set("props.border.bottom", value);
		}else if("none".equals(property)){
			update.set("props.border.left" , value);
			update.set("props.border.top", value);
			update.set("props.border.right" , value);
			update.set("props.border.bottom", value);
		}else{
		    update.set("props.border." + property, value);
		}
		mongoTemplate.updateMulti(query, update, MCol.class, excelId);
		
	}


}
