package com.acmr.excel.aop;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

import com.acmr.excel.model.RowCol;

public class Test {
	
	public void f(List<RowCol> a,Object b){
		System.out.println(a.get(0));
		System.out.println();
		
	}
	
	public static void main(String[] args) throws NoSuchMethodException, SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException {
        int a=1;
        String type = int.class.getName();
		int length= type.lastIndexOf(".");
		type =type.substring(length+1);
		System.out.println(type);
		
        Test test = new Test();
        List<RowCol> al = new ArrayList<RowCol>();
        RowCol rc1 = new RowCol();
        rc1.setAlias("6");
        al.add(rc1);
        List<Object> bl = new ArrayList<Object>();
        bl.add(2);
        RowCol rc = new RowCol();
        rc.setAlias("5");
        
        Object[] arg = new Object[]{"a","b"};
       // Execute.exec(test.getClass().getName(), "n", arg);
        Execute.getMethodParamTypes(test, "n");
        bl.remove(0);
        System.out.println(bl.size());
        bl.add("a");
        System.out.println(bl.size());
        
 /*       Class<?>[] typeArg = new Class<?>[arg.length];
       
        for(int i=0;i<arg.length;i++){
        	if(arg[i]==null){
        		typeArg[i]=Object.class;
        	}else if(arg[i].getClass()==ArrayList.class){
        		typeArg[i]=List.class;
        	}else{
        		typeArg[i]=Object.class;
        	}
        	
        }
  
 
	Method method = test.getClass().getDeclaredMethod("n",typeArg);
	method.invoke(test, arg);*/
		
		/*Function fn = Function.identity();
        Object mothod = fn.apply(test);
       */
     /*   String str = "12";
        List<Object> al = new ArrayList<Object>();
        
        if(str.getClass().equals(String.class) ){
        	System.out.println(1);
        }
        
        List<String> list = new ArrayList<>();
        list.add("1");
        list.add("2");
        */
        
	}
	
	public void n(String a,Object b){
		System.out.println(a);
		System.out.println(b);
		
	}
	
	public void n(String a,Collection<? extends Object> object){
		System.out.println(a);
		System.out.println("");
		
	}
	
}
