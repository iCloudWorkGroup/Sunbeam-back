package test;

import java.util.ArrayList;
import java.util.List;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;

import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.mongo.MRowColList;

public class RowColTest {
 public static void main(String[] args) {
	
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
	    Query query = new Query();
      /*		query.addCriteria(Criteria.where("_id").is("cList").and("rcList.alias").is("1"));
		Update update = new Update();
		update.set("rcList.length", 120);
		mongoTemplate.updateFirst(query, update, "6d16ae1b-7191-4c4d-b67b-a49798a22d00");
		*/
	/*  query.addCriteria(Criteria.where("id").is("aa").and("aa.sort").is(4));
	    // query.addCriteria(Criteria.where("aa.sort").is(1.0));
	    List<Object> list = new ArrayList<Object>();
	    list.add(3);
	    list.add(4);
	    Update update = new Update();
	  
	  // update.addToSet("aa", sort);
	    //update.pushAll("comments", list.toArray());
	   // update.push("aa", sort);
	    mongoTemplate.updateFirst(query, update, "aa");
	   // mongoTemplate.insert(list,"aa");
	    //  List<Sort> sort1 =  mongoTemplate.find(query, Sort.class,"aa");
	    System.out.println();
	    */
	    
	 /*   query.addCriteria(Criteria.where("_id").is("cList").and("rcList.alias")
				.is("5").and("sheetId").is("13300b0d-238c-42d5-8ca3-36d38f4b8d790"));
	    MRowColList rc = mongoTemplate.findOne(query, MRowColList.class, "13300b0d-238c-42d5-8ca3-36d38f4b8d79");
	    System.out.println();*/
	    Sort sort = new Sort();
	    sort.setSort(5);
	    sort.setId(1);
	    Sort sort1 = new Sort();
	    sort1.setSort(3);
	    sort1.setId(3);
	    List<Sort> list = new ArrayList<Sort>();
	    list.add(sort);
	    list.add(sort1);
	    //mongoTemplate.insert(list, "aa");
	    mongoTemplate.remove(list, "aa");
	    //mongoTemplate.remove(sort, "aa");
}
} 
