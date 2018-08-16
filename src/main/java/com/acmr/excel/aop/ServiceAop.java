package com.acmr.excel.aop;

import org.apache.logging.log4j.core.config.Order;
import org.aspectj.lang.annotation.Aspect;
import org.aspectj.lang.annotation.Before;
import org.aspectj.lang.annotation.Pointcut;
import org.springframework.stereotype.Component;


public class ServiceAop {

	@Pointcut("execution(* com.acmr.excel.service.*.*(..))")
	public void pointCut2() {
	}
	
	@Before("pointCut2()")
	public void beforeService(){
		System.out.println("进入service层");
	}
}
