package test;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;

import com.acmr.excel.model.mongo.MExcelRow;

public class RowStatusTest {
	
	public static void main(String[] args) {
		//ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		//MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
		ApplicationContext ac = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		MongoTemplate mongoTemplate = (MongoTemplate)ac.getBean("mongoTemplate");
		Query query = new Query();
		query.addCriteria(Criteria.where("excelRow.code").is("2"));
		Update update = new Update();
		update.set("excelRow.rowhidden", false);
		mongoTemplate.updateFirst(query, update,MExcelRow.class, "d02b0e50-02f0-4c7e-b2ac-486e6e487e7c");
	}

}
