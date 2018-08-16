package com.acmr.excel.dao.impl;

import java.io.Serializable;
import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MRowColCellDao;
import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.mongo.MRowColCell;

@Repository("mrowColCellDao")
public class MRowColCellDaoImpl implements MRowColCellDao,Serializable {

	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			String alias, String type) {

		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where(type).is(alias).and("sheetId")
				.is(sheetId).and("_class").is(MRowColCell.class.getName()));
		List<MRowColCell> list = mongoTemplate.find(new Query(criatira),
				MRowColCell.class, excelId);
		return list;
	}

	@Override
	public List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			List<String> aliasList, String type) {

		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where(type).in(aliasList).and("sheetId")
				.is(sheetId).and("_class").is(MRowColCell.class.getName()));
		List<MRowColCell> list = mongoTemplate.find(new Query(criatira),
				MRowColCell.class, excelId);
		return list;
	}

	@Override
	public List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			List<String> rowList, List<String> colList) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_class")
				.is(MRowColCell.class.getName()).and("sheetId").is(sheetId)
				.and("row").in(rowList).and("col").in(colList));

		List<MRowColCell> list = mongoTemplate.find(query, MRowColCell.class,
				excelId);

		return list;
	}
	
	@Override
	public List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			List<String> cellIdList) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_class")
				.is(MRowColCell.class.getName()).and("sheetId").is(sheetId)
				.and("cellId").in(cellIdList));

		List<MRowColCell> list = mongoTemplate.find(query, MRowColCell.class,
				excelId);

		return list;
	}
	
	@Override
	public List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			String cellId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_class")
				.is(MRowColCell.class.getName()).and("sheetId").is(sheetId)
				.and("cellId").is(cellId));

		List<MRowColCell> list = mongoTemplate.find(query, MRowColCell.class,
				excelId);

		return list;
	}

	@Override
	public void delMRowColCell(String excelId, String sheetId, String field,
			String alias) {

		mongoTemplate.remove(new Query(Criteria.where(field).is(alias)
				.and("_class").is(MRowColCell.class.getName()).and("sheetId")
				.is(sheetId)), MRowColCell.class, excelId);

	}

	@Override
	public void delMRowColCellList(String excelId, String sheetId,
			List<String> rowList, List<String> colList) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_class")
				.is(MRowColCell.class.getName()).and("sheetId").is(sheetId)
				.and("row").in(rowList).and("col").in(colList));
		mongoTemplate.remove(query, MRowColCell.class, excelId);
	}

	@Override
	public void delMRowColCellList1(String excelId, String sheetId,
			List<String> cellIdList) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_class")
				.is(MRowColCell.class.getName()).and("sheetId").is(sheetId)
				.and("cellId").in(cellIdList));
		mongoTemplate.remove(query, MRowColCell.class, excelId);

	}

	@Override
	public void updateMRowColCell(String excelId, String sheetId,
			String originalId, String modifiedId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("cellId").is(originalId).and("sheetId")
				.is(sheetId).and("_class").is(MRowColCell.class.getName()));
		Update update = new Update();
		update.set("cellId", modifiedId);
		mongoTemplate.updateMulti(query, update, MRowColCell.class, excelId);

	}

}
