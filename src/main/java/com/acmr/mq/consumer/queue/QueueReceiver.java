/**
 * 
 */
package com.acmr.mq.consumer.queue;

import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import javax.annotation.Resource;
import javax.jms.JMSException;
import javax.jms.Message;
import javax.jms.MessageListener;
import javax.jms.ObjectMessage;
import javax.jms.TextMessage;

import org.apache.log4j.Logger;
import org.springframework.stereotype.Service;


import com.acmr.excel.service.MCellService;
import com.acmr.excel.service.MColService;
import com.acmr.excel.service.MExcelService;
import com.acmr.excel.service.MRowService;
import com.acmr.excel.service.MSheetService;

import com.acmr.excel.service.impl.MongodbServiceImpl;
import com.acmr.mq.Model;




/**
 * @描述 队列消息监听器
 */
@Service
public class QueueReceiver implements MessageListener {
	private static Logger logger = Logger.getLogger(QueueReceiver.class);
	
	@Resource
	private MongodbServiceImpl mongodbServiceImpl;
	
	@Resource
	private MExcelService mexcelService;
	@Resource
	private MCellService mcellService;
	@Resource
	private MRowService mrowService;
	@Resource
	private MColService mcolService;
	@Resource
	private MSheetService msheetService;

	@Override
	public synchronized void onMessage(Message message) {
		if (message instanceof ObjectMessage) {
			ObjectMessage objectMessage = (ObjectMessage) message;
			try {
				Model model = (Model) objectMessage.getObject();
				String excelId = model.getExcelId();
				int step = model.getStep();
				logger.info("**********receive message excelId : "+excelId + " === step : " + step + "== reqPath : "+ model.getReqPath());
				ExecutorService executor = Executors.newFixedThreadPool(1);
				
				Runnable worker = new WorkerThread2(step,excelId,model, mcellService, mexcelService,mrowService,mcolService,msheetService);
				executor.execute(worker);
				executor.shutdown();
			} catch (JMSException e) {
				logger.info(e.getLocalizedMessage());
				logger.info(e.getMessage());
				e.printStackTrace();
			}
		} else if (message instanceof TextMessage) {
			try {
				System.out.println(((TextMessage) message).getText());
			} catch (JMSException e) {
				e.printStackTrace();
			}
		}else{
			System.out.println("message没有被处理^^^^^^^");
		}
	}

	

}
