package com.acmr.excel.service.impl;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.service.MExcelService;

@Service("mexcelService")
public class MExcelServiceImpl implements MExcelService {

	@Resource
	private BaseDao baseDao;
	
	public int getStep(String id) {
		
		return baseDao.getStep(id);
	}

}
