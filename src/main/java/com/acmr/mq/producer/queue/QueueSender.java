/**
 * 
 */
package com.acmr.mq.producer.queue;

import java.util.UUID;

import javax.annotation.Resource;
import javax.jms.JMSException;
import javax.jms.Message;
import javax.jms.Session;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.jms.JmsException;
import org.springframework.jms.core.JmsTemplate;
import org.springframework.jms.core.MessageCreator;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Service;

import com.acmr.excel.model.Constant;
import com.acmr.mq.Model;

/**
 * @作者 Goofy
 * @邮件 252878950@qq.com
 * @日期 2014-4-1上午9:40:24
 * @描述 发送消息到队列
 */
@Service
public class QueueSender {
	
	@Autowired
	@Qualifier("jmsQueueTemplate")
	//@Resource(name="jmsQueueTemplate")
	private JmsTemplate jmsTemplate;//通过@Qualifier修饰符来注入对应的bean
	
	/**
	 * 发送一条消息到指定的队列（目标）
	 * @param queueName 队列名称
	 * @param message 消息内容
	 */
	public void send(String queueName,final Model model){
			//Thread.sleep(model.getStep()*100);
			jmsTemplate.send(Constant.queueName, new MessageCreator() {
				@Override
				public Message createMessage(Session session) throws JMSException {
					return session.createObjectMessage(model);
				}
			});
	}
	
}
