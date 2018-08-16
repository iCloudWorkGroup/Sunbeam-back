package test;

import java.util.ArrayList;
import java.util.List;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;

import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Cell;
import com.acmr.excel.service.impl.MCellServiceImpl;

public class AopTest {

	public static void main(String[] args) {
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext( 
				new String[]{"config/applicationContext-core.xml", "/config/mongodb.xml",
						"config/ActiveMQ.xml","config/spring-mvc.xml"
						});
		
		/*MRowColDaoImpl mrc = (MRowColDaoImpl) applicationContext.getBean("mrowColDao");
		List<RowCol> list = new ArrayList<RowCol>();
		mrc.getColList(list, "c37e5497-3935-4b8c-9c99-56c427b22185", "c37e5497-3935-4b8c-9c99-56c427b221850")*/;
		BaseDao  mrc = (BaseDao) applicationContext.getBean("baseDao");
		
        Cell cell = new Cell();	
        cell.setAlign("4");
        Cell cell2 = new Cell();	
        cell2.setAlign("5");
		//mrc.splitCell(cell, 1, "");
        List<Object> list = new ArrayList<Object>();
        list.add(cell);
        list.add(null);
        mrc.insertList("aa", list);
	}
}
