package com.acmr.excel.interceptor;

import java.util.List;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.web.servlet.HandlerInterceptor;
import org.springframework.web.servlet.ModelAndView;

import com.acmr.excel.model.Constant;

public class CorsInterceptor implements HandlerInterceptor {

	@Override
	public void afterCompletion(HttpServletRequest paramHttpServletRequest,
			HttpServletResponse paramHttpServletResponse, Object paramObject,
			Exception paramException) throws Exception {
		// System.out.println("afterCompletion");
	}

	@Override
	public void postHandle(HttpServletRequest paramHttpServletRequest,
			HttpServletResponse paramHttpServletResponse, Object paramObject,
			ModelAndView paramModelAndView) throws Exception {

	}

	@Override
	public boolean preHandle(HttpServletRequest paramHttpServletRequest,
			HttpServletResponse paramHttpServletResponse, Object paramObject)
			throws Exception {
		if (paramHttpServletRequest.getHeader("Access-Control-Request-Method") != null
				&& "OPTIONS".equals(paramHttpServletRequest.getMethod())) {
			String accessControlAllowOrign = paramHttpServletResponse
					.getHeader("Access-Control-Allow-Origin");
			List<String> accessControlAllowOriginList = Constant.accessControlAllowOriginList;
			if (accessControlAllowOriginList.contains(accessControlAllowOrign)) {
				paramHttpServletResponse.setHeader(
						"Access-Control-Allow-Origin", accessControlAllowOrign);
				// paramHttpServletResponse.setHeader("Cache-Control",
				// "Public");
				// paramHttpServletResponse.setHeader("Access-Control-Max-Age","1728000");//20天

			} else {
				paramHttpServletResponse.setHeader(
						"Access-Control-Allow-Origin", "http://localhost:8080");
			}
		}
//		String reqUri = paramHttpServletRequest.getRequestURI();
		
		
        return true;
	}

}
