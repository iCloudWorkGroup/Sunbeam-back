package test;

import java.lang.reflect.Field;

import com.acmr.excel.model.complete.Content;

import acmr.excel.pojo.Constants.CELLTYPE;

public class ReflectTest {
   public static void main(String[] args) throws Exception {
	   Content content = new Content();
	   content.setColor("aa");
	   Class c = content.getClass();
		Content mc = new Content();
		Field[] cFields = c.getDeclaredFields();
		for(Field f:cFields){
			f.setAccessible(true);
			Object value = f.get(content);
			if(null!=value){
				Field mf = mc.getClass().getDeclaredField(f.getName());
				mf.setAccessible(true);
				mf.set(mc, value);
			}
	    }
	System.out.println(mc);	
	String a = CELLTYPE.BLANK.toString();
	System.out.println(a);
}
}
