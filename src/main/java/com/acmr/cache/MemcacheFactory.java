//package com.acmr.cache;
//
//import java.io.IOException;
//import java.net.InetSocketAddress;
//
//
//
//
//import com.acmr.excel.util.PropertiesReaderUtil;
//import com.acmr.excel.util.StringUtil;
//
//import net.spy.memcached.MemcachedClient;
//
//public enum MemcacheFactory {
//	CACHESOURCE;
//	private MemcachedClient memcachedClient;
//
//	private MemcacheFactory() {
//		String address = PropertiesReaderUtil.get("memcache.server");
//		String port = PropertiesReaderUtil.get("memcache.port");
//		int portt = 11211;
//		if(!StringUtil.isEmpty(port)){
//			portt = Integer.valueOf(port);
//		}
//		try {
//			memcachedClient = new MemcachedClient(new InetSocketAddress(address, portt));
//		} catch (IOException e) {
//			e.printStackTrace();
//		}
//	}
//
//	public MemcachedClient getMemcacheClient() {
//		return memcachedClient;
//	}
//
//}
