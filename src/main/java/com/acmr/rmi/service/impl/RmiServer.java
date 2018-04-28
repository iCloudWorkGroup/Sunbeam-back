package com.acmr.rmi.service.impl;

import java.rmi.AccessException;
import java.rmi.NoSuchObjectException;
import java.rmi.NotBoundException;
import java.rmi.Remote;
import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.registry.Registry;
import java.rmi.server.UnicastRemoteObject;

import javax.naming.Context;
import javax.naming.InitialContext;
import javax.servlet.ServletContextEvent;
import javax.servlet.ServletContextListener;










import net.spy.memcached.MemcachedClient;

import org.springframework.web.context.WebApplicationContext;
import org.springframework.web.context.support.WebApplicationContextUtils;

import com.acmr.excel.util.PropertiesReaderUtil;
import com.acmr.rmi.service.RmiService;

public class RmiServer implements ServletContextListener {
	private Registry reg;
	@Override
	public void contextDestroyed(ServletContextEvent event) {
//		WebApplicationContext springContext = WebApplicationContextUtils.getWebApplicationContext(event.getServletContext());
//		ConfigurableApplicationContext configurableApplicationContext = (ConfigurableApplicationContext)springContext;
//		BeanDefinitionRegistry beanDefinitionRegistry = (DefaultListableBeanFactory)configurableApplicationContext.getBeanFactory();
//		beanDefinitionRegistry.removeBeanDefinition("rmiServiceImpl");
			try {
				String[] regNames = reg.list();
				for (String lName : regNames) {
				    Remote lRemoteObj = reg.lookup(lName);
				    reg.unbind(lName);
				    UnicastRemoteObject.unexportObject(lRemoteObj, true);
				}
				UnicastRemoteObject.unexportObject(reg, true);
				
			} catch (AccessException e) {
				e.printStackTrace();
			} catch (NoSuchObjectException e) {
				e.printStackTrace();
			} catch (RemoteException e) {
				e.printStackTrace();
			} catch (NotBoundException e) {
				e.printStackTrace();
			}
	}

	@Override
	public void contextInitialized(ServletContextEvent event) {
//		WebApplicationContext springContext = WebApplicationContextUtils.getWebApplicationContext(event.getServletContext());
//		MemcachedClient memcachedClient = (MemcachedClient) springContext.getBean("memcachedClient");
//		try {
//			RmiService rmi = new RmiServiceImpl(memcachedClient);
//			String port = PropertiesReaderUtil.get("rmi.port");
//			reg = LocateRegistry.createRegistry(Integer.valueOf(port));
//			reg.rebind("RmiService", rmi);
//			System.out.println("rmi server start");
//		} catch (Exception e) {
//			e.printStackTrace();
//		}

	}
}
