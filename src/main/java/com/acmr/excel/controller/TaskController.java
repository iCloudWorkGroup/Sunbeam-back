package com.acmr.excel.controller;

import javax.annotation.Resource;

import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Controller;

import com.acmr.excel.service.MBookService;


public class TaskController {

	@Resource
	private MBookService mbookService;
	
	
	@Scheduled(cron = "*/10 * * * * ?")
	public void delJob(){
		System.out.println("执行定时任务");
		mbookService.delCollections();
	}
}
