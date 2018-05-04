package test;

import java.util.ArrayList;
import java.util.List;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;

import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.mongo.MExcelCell;

public class Test1 {

	public static void main(String[] args) {
		
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
		//查找关系映射表
		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("row").gte(1).lte(8),
						Criteria.where("col").gte(1).lte(8));
		List<RowColCell> list = mongoTemplate.find(new Query(criatira), RowColCell.class,"d592f68c-0c0a-481f-b7a6-728262715e41");
				//查找对应的cell
		List<String>  inlist = new ArrayList<String>();
		for(RowColCell rcc:list){
					inlist.add(rcc.getCellId());
			}
		List<MExcelCell> cellList = mongoTemplate.find(new Query(Criteria.where("_id").in(inlist)), MExcelCell.class,"d592f68c-0c0a-481f-b7a6-728262715e41");
		System.out.println(cellList.size());
		/*List<String> list = new ArrayList<String>();
		list.set(2, "a");*/
	/*	List<RowCol> list = new ArrayList<RowCol>();
		RowCol rc = new RowCol();
		rc.setAlias("1");
		list.add(rc);
		RowCol rc1 = new RowCol();
		rc1.setAlias("1");*/
		
		
		
	}

}
