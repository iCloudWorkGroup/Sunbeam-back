package com.acmr.rmi.service;

import java.rmi.AccessException;
import java.rmi.NotBoundException;
import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.registry.Registry;
import java.rmi.server.ServerNotActiveException;

import acmr.excel.ExcelException;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelSheet;


public class Test {
   public static void main(String[] args) throws Exception {
	   Registry registry = LocateRegistry.getRegistry("192.168.3.84", 11999);
	   RmiService rmiService = (RmiService) registry.lookup("RmiService");
	   ExcelBook book1 = new ExcelBook();
	   /*book1.LoadExcel("D://123.xlsx");
	   ExcelSheet  sheet = new ExcelSheet();
	   sheet.addRow();
	   sheet.addColumn();
	   book1.getSheets().add(sheet);
	   rmiService.saveExcelBook("123", book1);
	   rmiService.saveExcelBook("123init", book1);
	   rmiService.saveExcelBook("12345456", book1);*/
	   ExcelBook  book= rmiService.getExcelBook("1234aaaa", 0);
	  // rmiService.saveExcelBook("1234aaaa", book);
	   
	   try {
		   book.saveExcel("D://456.xlsx");
		} catch (ExcelException e) {
			e.printStackTrace();
		}
	   
	 //  book1.LoadExcel("D://789.xlsx");
	   
	  // rmiService.saveExcelBook("c33b55e4-43bf-4f09-8b08-5384fa65633d", book);
	   System.out.println(book);
}
}
