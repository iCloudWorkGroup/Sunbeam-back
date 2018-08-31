package com.acmr.excel.dao.impl;

import java.io.Serializable;
import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.model.mongo.MCell;

@Repository("mcellDao")
public class MCellDaoImpl implements MCellDao, Serializable {

	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<MCell> getMCellList(String excelId, String sheetId,
			List<String> cellIdList) {
		
	/*	Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("_id").in(cellIdList).and("sheetId")
				.is(sheetId));
		Query query = new Query(criatira).with(new Sort(Direction.DESC, "collectDate"));*/
		List<MCell> list = mongoTemplate.find(new Query(Criteria.where("_id")
				.in(cellIdList).and("sheetId").is(sheetId)), MCell.class,
				excelId);
		return list;
	}

	@Override
	public MCell getMCell(String excelId, String sheetId, String id) {

		MCell mcell = mongoTemplate.findOne(new Query(
				Criteria.where("_id").is(id).and("sheetId").is(sheetId)),
				MCell.class, excelId);
		return mcell;
	}

	@Override
	public void delMCell(String excelId, String sheetId,
			List<String> cellIdList) {

		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("_id").in(cellIdList).and("sheetId")
				.is(sheetId).and("_class").is(MCell.class.getName()));
		mongoTemplate.remove(new Query(criatira), MCell.class, excelId);

	}

	@Override
	public void updateContent(String property1, Object value1, String property2,
			Object value2, List<String> idList, String excelId,
			String sheetId) {
		Query query = new Query();
		query.addCriteria(
				Criteria.where("_id").in(idList).and("sheetId").is(sheetId));
		Update update = new Update();
		if (("displayTexts".equals(property1))) {
			update.set("content.displayTexts", value1);
			update.set("content.texts", value1);
		} else if ("type".equals(property2)) {
			update.set("content." + property1, value1);
			update.set("content." + property2, value2);
		} else {
			update.set("content." + property1, value1);
		}

		mongoTemplate.updateMulti(query, update, MCell.class, excelId);
	}

	@Override
	public void updateBorder(String property, Object value, List<String> idList,
			String excelId, String sheetId) {
		Query query = new Query();
		query.addCriteria(
				Criteria.where("_id").in(idList).and("sheetId").is(sheetId));
		Update update = new Update();
		if ("all".equals(property)) {
			update.set("border.left", value);
			update.set("border.top", value);
			update.set("border.right", value);
			update.set("border.bottom", value);
		} else if ("none".equals(property)) {
			update.set("border.left", value);
			update.set("border.top", value);
			update.set("border.right", value);
			update.set("border.bottom", value);
		} else {
			update.set("border." + property, value);
		}
		mongoTemplate.updateMulti(query, update, MCell.class, excelId);

	}

	@Override
	public void updateCustomProp(String property, Object value,
			List<String> idList, String excelId, String sheetId) {
		Query query = new Query();
		query.addCriteria(
				Criteria.where("_id").in(idList).and("sheetId").is(sheetId));
		Update update = new Update();
		update.set("customProp." + property, value);
		mongoTemplate.updateMulti(query, update, MCell.class, excelId);
	}

}
