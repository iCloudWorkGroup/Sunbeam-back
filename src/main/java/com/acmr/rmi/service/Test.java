package com.acmr.rmi.service;

import java.rmi.AccessException;
import java.rmi.NotBoundException;
import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.registry.Registry;


public class Test {
   public static void main(String[] args) throws AccessException, RemoteException, NotBoundException {
	   Registry registry = LocateRegistry.getRegistry("192.168.3.84", 11999);
	   RmiService rmiService = (RmiService) registry.lookup("RmiService");
	   rmiService.getExcelBook("f979939b-9f1c-4e83-bbac-5063b0e646b4", 0);
}
}
