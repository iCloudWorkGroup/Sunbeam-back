package test;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;

import com.acmr.excel.model.RowColCell;

public class RowColCellTest {
	public static void main(String[] args) {
		
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
		Query query = new Query();
		query.addCriteria(Criteria.where("cellId").is("bb"));
		Update update = new Update();
		update.set("cellId", "1_2");
		mongoTemplate.updateMulti(query, update, RowColCell.class,"d02b0e50-02f0-4c7e-b2ac-486e6e487e7c");
		//mongoTemplate.remove(new Query(Criteria.where("row").is(17)),RowColCell.class, "d02b0e50-02f0-4c7e-b2ac-486e6e487e7c");
	}

}
