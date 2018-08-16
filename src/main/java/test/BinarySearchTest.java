package test;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.junit.BeforeClass;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import org.springframework.context.ApplicationContext;

import com.acmr.excel.dao.impl.MRowColDaoImpl;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.util.BinarySearch;

@RunWith(Parameterized.class) 
public class BinarySearchTest {
	
	private static List<RowCol> list;
	
	private int top;
	
	private int index;
	
	public BinarySearchTest(int top,int index){
		this.top = top;
		this.index = index;
	}
	
	@Parameters
	public static Collection data(){
		return Arrays.asList(new Object[][]{
			{-2,0},
			{0,0},
			{720,10},
			{1008,14},
			{1600,14}
		});
		
	}

	@BeforeClass
	public static void setUp() throws Exception {

		@SuppressWarnings("resource")
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext( 
				new String[]{"config/applicationContext-core.xml", "/config/mongodb.xml",
						"config/ActiveMQ.xml","config/spring-mvc.xml"
						});
		
		MRowColDaoImpl mrc = (MRowColDaoImpl) applicationContext.getBean("mrowColDao");
		list = new ArrayList<RowCol>();
		mrc.getColList(list, "c37e5497-3935-4b8c-9c99-56c427b22185", "c37e5497-3935-4b8c-9c99-56c427b221850");
		
	}
	

	@Test
	public void testBinarySearch() {
		assertEquals(index, BinarySearch.binarySearch(list, top));
	}

}
