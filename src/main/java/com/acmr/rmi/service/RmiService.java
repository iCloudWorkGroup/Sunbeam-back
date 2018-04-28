package com.acmr.rmi.service;

import java.rmi.Remote;
import java.rmi.RemoteException;

import acmr.excel.pojo.ExcelBook;

public interface RmiService extends Remote {
	/**
	 * 保存excel
	 * @param excelId
	 * @param excelBook
	 */
	public boolean saveExcelBook(String excelId,ExcelBook excelBook) throws RemoteException;
	/**
	 * 获取excel
	 */
	public ExcelBook getExcelBook(String excelId,int step) throws RemoteException;
}
