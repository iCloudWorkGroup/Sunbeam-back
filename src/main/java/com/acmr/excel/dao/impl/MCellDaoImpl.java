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
import com.acmr.excel.model.mongo.MExcelCell;

@Repository("mcellDao")
public class MCellDaoImpl implements MCellDao {
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<RowColCell> getRowColCellList(String excelId,Integer row) {
		
		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("row").is(row));
		List<RowColCell> list = mongoTemplate.find(new Query(criatira), RowColCell.class,excelId);
		return list;
	}
	
	@Override
	public void delRowColCell(String excelId, String field,Integer alias) {
		
		mongoTemplate.remove(new Query(Criteria.where(field).is(alias)),RowColCell.class, excelId);
		
	}

	@Override
	public void delRowColCell(Integer col, String excelId) {
		
		mongoTemplate.remove(new Query(Criteria.where("col").is(col)),RowColCell.class, excelId);
		
	}
	
	@Override
	public void updateRowColCell(String excelId, String originalId, String modifiedId) {
		Query query = new Query();
		query.addCriteria(Criteria.where("cellId").is(originalId));
		Update update = new Update();
		update.set("cellId", modifiedId);
		mongoTemplate.updateMulti(query, update, RowColCell.class,excelId);
		
	}

	@Override
	public List<MExcelCell> getMCellList(String excelId, List<String> cellIdList) {
		
		List<MExcelCell> list = mongoTemplate.find(new Query(Criteria.where("_id").in(cellIdList)), MExcelCell.class,excelId);
		return list;
	}
	
	@Override
	public MExcelCell getMCell(String excelId, String id) {
		
		MExcelCell mexcelCell =  mongoTemplate.findOne(new Query(Criteria.where("_id").is(id)), MExcelCell.class,excelId);
		return mexcelCell;
	}

	@Override
	public List<RowColCell> getRowColCellList(Integer col, String excelId) {
		
		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("col").is(col));
		List<RowColCell> list = mongoTemplate.find(new Query(criatira), RowColCell.class,excelId);
		return list;
	}

	@Override
	public void delCell(String excelId, List<String> cellIdList) {
		
		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("_id").in(cellIdList));
		mongoTemplate.remove(new Query(criatira), MExcelCell.class,excelId);
		
	}

	@Override
	public void updateFont(String name,Object property,List<String> idList,String excelId) {
		Query query = new Query();
	    query.addCriteria(Criteria.where("_id").in(idList));
	    Update update = new Update();
	    update.set("excelCell.cellstyle.font."+name, property);
	    mongoTemplate.updateMulti(query, update, MExcelCell.class,excelId);
	}


	@Override
	public void updateHidden(String type, String alias,boolean state, String excelId) {
		Query query = new Query();
	    query.addCriteria(Criteria.where(type).is(alias));
	    Update update = new Update();
	    update.set("excelCell.cellstyle.hidden", state);
	    mongoTemplate.updateFirst(query, update, MExcelCell.class,excelId);
		
	}

}
