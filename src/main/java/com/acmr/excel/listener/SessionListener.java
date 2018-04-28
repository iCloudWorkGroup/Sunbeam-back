package com.acmr.excel.listener;

import java.util.Enumeration;

import javax.servlet.http.HttpSession;
import javax.servlet.http.HttpSessionEvent;
import javax.servlet.http.HttpSessionListener;

import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;
import org.springframework.context.support.FileSystemXmlApplicationContext;

import com.acmr.excel.dao.ExcelDao;
import com.acmr.excel.model.OnlineExcel;
import com.acmr.excel.model.complete.CompleteExcel;
import com.alibaba.fastjson.JSON;
/**
 * session监听类
 * @author jinhr
 *
 */

public class SessionListener implements HttpSessionListener {
	/**
	 * session创建时监听
	 * @param event HttpSessionEvent对象
	 */
	@Override
	public void sessionCreated(HttpSessionEvent event) {
	}
	/**
	 * session失效是监听
	 * @param sessionEvent HttpSessionEvent对象
	 */
	
	@Override
	public void sessionDestroyed(HttpSessionEvent sessionEvent) {
//		System.out.println("session destory");
//		ApplicationContext ac = new ClassPathXmlApplicationContext("com/acmr/excel/excel-service.xml"); 
//		ExcelDao excelDao = (ExcelDao)ac.getBean("excelDao");
//		HttpSession session = sessionEvent.getSession();
//		Enumeration<String> e = session.getAttributeNames();
//		while (e.hasMoreElements()) {
//            String s = e.nextElement();
//            System.out.println(s + " == " + session.getAttribute(s));
//            CompleteExcel excel = (CompleteExcel)session.getAttribute(s);
//            OnlineExcel oExcel = new OnlineExcel();
//            oExcel.setExcelId(s);
//            oExcel.setJsonObject(JSON.toJSONString(excel));
//            try {
//				excelDao.saveExcel(oExcel);
//			} catch (Exception e1) {
//				e1.printStackTrace();
//			}
//        }
		
	}

}
