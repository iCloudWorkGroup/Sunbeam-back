package com.acmr.rmi.service;

import java.rmi.Remote;
import java.rmi.RemoteException;
import java.rmi.server.ServerNotActiveException;

import acmr.excel.pojo.ExcelBook;

public interface RmiService extends Remote {
	/**
	 * 保存excel
	 * @param excelId
	 * @param excelBook
	 */
	public boolean saveExcelBook(String excelId,ExcelBook excelBook) throws RemoteException,ServerNotActiveException;
	/**
	 * 获取excel
	 */
	public ExcelBook getExcelBook(String excelId,int step) throws RemoteException,ServerNotActiveException;
}
