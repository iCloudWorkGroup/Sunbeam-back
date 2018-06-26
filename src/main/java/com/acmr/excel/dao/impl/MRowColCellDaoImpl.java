package com.acmr.excel.dao.impl;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MRowColCellDao;
import com.acmr.excel.model.mongo.MRowColCell;

@Repository("mrowColCellDao")
public class MRowColCellDaoImpl implements MRowColCellDao {

	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			List<String> rowList, List<String> colList) {

		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("row").in(rowList),
				Criteria.where("col").in(colList),
				Criteria.where("sheetId").in(sheetId));
		List<MRowColCell> list = mongoTemplate.find(new Query(criatira),
				MRowColCell.class, excelId);
		return list;
	}

}
