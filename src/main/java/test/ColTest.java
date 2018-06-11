package test;

import java.util.ArrayList;
import java.util.List;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;

import com.acmr.excel.model.mongo.MRow;

public class ColTest {
  public static void main(String[] args) {
	  ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
	 MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
	 List<String> rowList = new ArrayList<String>();
	 rowList.add("1");
	 rowList.add("2");
	 Query query = new Query();
	 query.addCriteria(Criteria.where("alias").in(rowList));
	 List<MRow> mRowList = new ArrayList<>();
	  mRowList = mongoTemplate.find(query, MRow.class, "0f2b8ccd-5841-48a8-8da5-c5b984e4d9cf");
	  System.out.println(MRow.class.getName());
	  System.out.println(mRowList.size());
}
}
