package com.acmr.excel.aop;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class Execute {

	public static void exec(Object object, String methodName, Object[] arg) {
		try {
	/*	Class<?> cla = Class.forName(object); 
		Object  ob = cla.newInstance();*/
		Class<?>[] typeArg = getMethodParamTypes(object,methodName);
		
			Method method = object.getClass().getDeclaredMethod(methodName,
					typeArg);
			method.invoke(object, arg);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static Class<?>[] getMethodParamTypes(Object object,
			String methodName)  {
		Class<?>[] paramTypes = null;
		try {
			Method[] methods = object.getClass().getMethods();// 全部方法
			for (int i = 0; i < methods.length; i++) {
				if (methodName.equals(methods[i].getName())) {// 和传入方法名匹配
					 paramTypes = methods[i].getParameterTypes();
				/*	paramTypes = new Class[params.length];
					for (int j = 0; j < params.length; j++) {
						paramTypes[j] = Class.forName(params[j].getName());
					}*/
					break;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return paramTypes;
	}

}
