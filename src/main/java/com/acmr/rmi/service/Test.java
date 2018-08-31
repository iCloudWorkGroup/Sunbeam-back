package com.acmr.rmi.service;

import java.rmi.AccessException;
import java.rmi.NotBoundException;
import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.registry.Registry;
import java.rmi.server.ServerNotActiveException;


public class Test {
   public static void main(String[] args) throws AccessException, RemoteException, NotBoundException, ServerNotActiveException {
	   Registry registry = LocateRegistry.getRegistry("192.168.3.84", 11999);
	   RmiService rmiService = (RmiService) registry.lookup("RmiService");
	   rmiService.getExcelBook("ebc482ea-2e12-4f0f-9f9a-a9dbf187a321", 0);
}
}
