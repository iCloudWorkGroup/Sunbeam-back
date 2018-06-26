package com.acmr.excel.dao.impl;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MRowColCell;

@Repository("mcellDao")
public class MCellDaoImpl implements MCellDao {

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
	public void delMRowColCell(String excelId, String sheetId, String field,
			String alias) {

		mongoTemplate.remove(new Query(Criteria.where(field).is(alias)
				.and("_class").is(MRowColCell.class.getName()).and("sheetId")
				.is(sheetId)), MRowColCell.class, excelId);

	}

	@Override
	public void delMRowColCell(Integer col, String excelId) {

		mongoTemplate.remove(new Query(Criteria.where("col").is(col)),
				RowColCell.class, excelId);

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
	public void delMRowColCellList(String excelId, String sheetId,
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

	@Override
	public List<MCell> getMCellList(String excelId, String sheetId,
			List<String> cellIdList) {

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
	public void updateHidden(String type, String alias, boolean state,
			String excelId) {
		Query query = new Query();
		query.addCriteria(Criteria.where(type).is(alias));
		Update update = new Update();
		update.set("excelCell.cellstyle.hidden", state);
		mongoTemplate.updateFirst(query, update, MExcelCell.class, excelId);

	}

	@Override
	public void updateContent(String property, Object content,
			List<String> idList, String excelId, String sheetId) {
		Query query = new Query();
		query.addCriteria(
				Criteria.where("_id").in(idList).and("sheetId").is(sheetId));
		Update update = new Update();
		if ("displayTexts".equals(property)) {
			update.set("content.displayTexts", null);
			update.set("content.texts", null);
		} else {
			update.set("content." + property, content);
		}

		mongoTemplate.updateMulti(query, update, MCell.class, excelId);
	}

	@Override
	public void updateSpan(List<String> cellIdList, String excelId,
			String sheetId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").in(cellIdList).and("sheetId")
				.is(sheetId));
		Update update = new Update();
		update.set("rowspan", 1);
		update.set("colspan", 1);
		mongoTemplate.updateMulti(query, update, MCell.class, excelId);
	}

}
