/*package com.acmr.excel.service.impl;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.model.Constant;
import com.acmr.excel.service.StoreService;
import com.whalin.MemCached.MemCachedClient;

@Service
public class MemJCServiceImpl implements StoreService {
	@Resource
	private MemCachedClient memCachedClient;

	@Override
	public Object get(String id) {
		return memCachedClient.get(id);
	}

	@Override
	public boolean set(String id, Object object) {
		boolean result = memCachedClient.set(id, object,Integer.valueOf(Constant.MEMCACHED_EXP_TIME));
		return result;
	}

}
*/