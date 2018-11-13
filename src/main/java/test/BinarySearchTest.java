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
			{1,0},
			{1234,14}
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
		mrc.getRowList(list, "048f5332-70b0-4e2e-811a-13ca53d86601init", "048f5332-70b0-4e2e-811a-13ca53d86601init0");
		
	}
	

	@Test
	public void testBinarySearch() {
		int xiabiao = BinarySearch.binarySearch(list, top);
		assertEquals(index,xiabiao );
	}

}
