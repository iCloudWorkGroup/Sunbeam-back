package com.acmr.excel.interceptor;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.web.method.HandlerMethod;
import org.springframework.web.servlet.HandlerInterceptor;
import org.springframework.web.servlet.ModelAndView;

import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.util.StringUtil;
import com.alibaba.fastjson.JSON;

public class ProtectInterceptor implements HandlerInterceptor{
	
	@Resource
	private MSheetDao msheetDao;

	@Override
	public boolean preHandle(HttpServletRequest req,
			HttpServletResponse res, Object handler) throws Exception {
		// 
		// System.out.println(handlerMethod.getMethod().getName());
		if(handler instanceof HandlerMethod) {
		 HandlerMethod handlerMethod = (HandlerMethod)handler;	
		 String methodName = handlerMethod.getMethod().getName();
		 System.out.println("进入拦截器:"+handlerMethod.getMethod().getName());
		 String excelId = req.getHeader("X-Book-Id");
		 String sheetId = excelId+0;
		 if((null != Constant.unInterceptAction.get(methodName))||(null==excelId)){
				return true;
		   }else{
				MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
				 if(null !=msheet.getProtect()&&msheet.getProtect()){
					 return false;
				 }else{
					 return true;
				 }
			 }
		 
	   }else{
		  return true;
	   }
	}

	@Override
	public void postHandle(HttpServletRequest request,
			HttpServletResponse response, Object handler,
			ModelAndView modelAndView) throws Exception {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void afterCompletion(HttpServletRequest request,
			HttpServletResponse response, Object handler, Exception ex)
					throws Exception {
		// TODO Auto-generated method stub
		
	}
	
	
	/**
	 * 把json串转化为实际对象
	 * 
	 * @param request
	 *            request对象
	 * @param t
	 *            泛型对象
	 * @return 实际对象
	 * @throws IOException
	 */

	protected <T> T getJsonDataParameter(HttpServletRequest request,
			Class<T> clazz) throws IOException {
		String body = null;
		StringBuilder stringBuilder = new StringBuilder();
		BufferedReader bufferedReader = null;
		try {
			InputStream inputStream = request.getInputStream();
			if (inputStream != null) {
				bufferedReader = new BufferedReader(new InputStreamReader(
						inputStream));
				char[] charBuffer = new char[128];
				int bytesRead = -1;
				while ((bytesRead = bufferedReader.read(charBuffer)) > 0) {
					stringBuilder.append(charBuffer, 0, bytesRead);
				}
			} else {
				stringBuilder.append("");
			}
		} catch (IOException ex) {
			throw ex;
		} finally {
			if (bufferedReader != null) {
				try {
					bufferedReader.close();
				} catch (IOException ex) {
					throw ex;
				}
			}
		}
		body = stringBuilder.toString();
		//System.out.println(body);
		if (StringUtil.isEmpty(body)) {
			try {
				return clazz.newInstance();
			} catch (InstantiationException | IllegalAccessException e) {
				e.printStackTrace();
			}
		}
		return JSON.parseObject(body, clazz);
	}


}
