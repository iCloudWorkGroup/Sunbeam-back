package com.acmr.excel.listener;

import javax.servlet.http.HttpSessionEvent;
import javax.servlet.http.HttpSessionListener;
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

	}

}
