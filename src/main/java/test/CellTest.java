package test;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;

import com.acmr.excel.model.mongo.MExcelCell;

import acmr.excel.pojo.ExcelCellStyle;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelFont;


public class CellTest {
public static void main(String[] args) {
	ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
	MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
   /* Query query = new Query();
    query.addCriteria(Criteria.where("_id").is("10_1"));
    Update update = new Update();
    ExcelCellStyle style = new ExcelCellStyle();
    ExcelFont font = style.getFont();
    font.setSize((short) 300);
    update.set("excelCell.cellstyle.font", font);
    mongoTemplate.updateFirst(query, update, MExcelCell.class,"9dad6428-2b6d-42dd-9e9a-185fdcfc0e12");*/
    MExcelCell mexcelCell =  mongoTemplate.findOne(new Query(Criteria.where("_id").is("10_1")), MExcelCell.class,"9dad6428-2b6d-42dd-9e9a-185fdcfc0e12");
    System.out.println(mexcelCell.getColId());
}
}
