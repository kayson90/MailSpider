package com.lenovo.tools.easmail;

public class EasApprovalTool {

    public static void main(String[] args) {

        MailReceiver mailReceiver = new ExchangeMailReceiver();
        //mailReceiver.configToReceiveMail();
        // start a thread to run the tool in background
        //UpdateTask updateTask = new UpdateTask();
        //updateTask.start();

    }

    /*class UpdateTask extends Thread {
    	public void run()  {
    		//
    		System.out.println("启动线程");
    		MailReceiver mailReceiver = new MailReceiver();
    	
    	

    	}
    }*/

}
