package com.acmr.rmi.service.impl;

import java.rmi.RemoteException;
import java.rmi.server.UnicastRemoteObject;

import javax.annotation.Resource;

import org.springframework.stereotype.Component;

import com.acmr.excel.service.MBookService;
import com.acmr.rmi.service.RmiService;

import acmr.excel.pojo.ExcelBook;


public class RmiServiceImpl extends UnicastRemoteObject implements RmiService {
	
	private MBookService mbookService;
	 
	protected RmiServiceImpl(MBookService mbookService) throws RemoteException {
		this.mbookService = mbookService;
	}

	

	@Override
	public boolean saveExcelBook(String excelId, ExcelBook excelBook) throws RemoteException {
		
		boolean result =  mbookService.saveExcelBook(excelBook, excelId);
	    return result;
	}

	@Override
	public ExcelBook getExcelBook(String excelId, int step) throws RemoteException {
		
		ExcelBook excelBook = mbookService.reloadExcelBook(excelId,step);
		
		return excelBook;
	}
	

}
