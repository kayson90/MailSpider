package com.lenovo.tools.easmail;


import microsoft.exchange.webservices.data.core.service.item.EmailMessage;



public interface MailReceiver{
	
	public void configToReceiveMail();
	//public void parseMessage();
	//public static String getFrom(EmailMessage ) throws MessagingException,ServiceLocalException ;
	//public static String getSentDate(EmailMessage,String )throws MessagingException, ServiceLocalException;
}