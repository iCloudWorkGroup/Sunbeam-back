package com.acmr.excel.distribute;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Distribute {
	
    private List<Target> container = new ArrayList<Target>();
    private Map<Object,Object> map = new HashMap<Object,Object>();
    
    public void add(Target target){
    	container.add(target);
    }
    
    public void exec(){
    	for(Target target :container ){
    		Object o = target.getObject();
    		Class<?> c = o.getClass();
    		try {
    			int type = target.getType();
    			if(type ==0){
    				Method method = c.getDeclaredMethod(target.getMethod());
					   method.invoke(o);
    			}else{
    				Method method = c.getDeclaredMethod(target.getMethod(), Map.class);
					   method.invoke(o, map);
    			}
				
			} catch (Exception e) {
				e.printStackTrace();
			}
    		
    		
    	}
    }
    
}
