package com.acmr.rmi.service;

import java.rmi.AccessException;
import java.rmi.NotBoundException;
import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.registry.Registry;
import java.rmi.server.ServerNotActiveException;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelSheet;


public class Test {
   public static void main(String[] args) throws AccessException, RemoteException, NotBoundException, ServerNotActiveException {
	   Registry registry = LocateRegistry.getRegistry("192.168.3.84", 11999);
	   RmiService rmiService = (RmiService) registry.lookup("RmiService");
	   ExcelBook book1 = new ExcelBook();
	   ExcelSheet  sheet = new ExcelSheet();
	   sheet.addRow();
	   sheet.addColumn();
	   book1.getSheets().add(sheet);
	   rmiService.saveExcelBook("123", book1);
	   rmiService.saveExcelBook("123init", book1);
	   ExcelBook  book= rmiService.getExcelBook("123", 0);
	   System.out.println(book);
}
}
