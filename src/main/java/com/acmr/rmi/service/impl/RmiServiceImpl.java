package com.acmr.rmi.service.impl;

import java.rmi.RemoteException;
import java.rmi.server.ServerNotActiveException;
import java.rmi.server.UnicastRemoteObject;

import org.apache.log4j.Logger;

import com.acmr.excel.service.MBookService;
import com.acmr.rmi.service.RmiService;

import acmr.excel.pojo.ExcelBook;


public class RmiServiceImpl extends UnicastRemoteObject implements RmiService {
	
	private static Logger logger = Logger.getLogger(RmiServiceImpl.class);
	
	private MBookService mbookService;
	 
	protected RmiServiceImpl(MBookService mbookService) throws RemoteException {
		this.mbookService = mbookService;
	}

	

	@Override
	public boolean saveExcelBook(String excelId, ExcelBook excelBook) throws RemoteException, ServerNotActiveException {
		
		String ip=UnicastRemoteObject.getClientHost();
	
		long start = System.currentTimeMillis();
		boolean result =  mbookService.saveExcelBook(excelBook, excelId);
		long end = System.currentTimeMillis();
		logger.info("客户端"+ip+";用时："+(end-start)+"ms");
	    return result;
	}

	@Override
	public ExcelBook getExcelBook(String excelId, int step) throws RemoteException, ServerNotActiveException {
		
		String ip=UnicastRemoteObject.getClientHost();
		
		long start = System.currentTimeMillis();
		ExcelBook excelBook = mbookService.reloadExcelBook(excelId,step);
		long end = System.currentTimeMillis();
		logger.info("客户端:"+ip+";用时："+(end-start)+"ms");
		return excelBook;
	}
	

}
