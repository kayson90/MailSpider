package com.lenovo.tools.easmail;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URI;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Timer;
import java.util.TimerTask;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

import com.sun.xml.internal.messaging.saaj.packaging.mime.MessagingException;

public class ExchangeMailReceiver implements MailReceiver {

    private String host = "mail.lenovo.com";
    private String username = "easquotation@lenovo.com";
    private String password = "xQaI-3962";

    // private Properties props;
    private Timer timer;

    ExchangeMailReceiver() {

        timer = new Timer();
        timer.schedule(new MailTask(), 0, 1 * 60 * 1000);// 每个1分钟调用1次MailTask类的run方法

        /*
         * 第一个参数是要操作的方法，第二个参数是要设定延迟的时间，第三个参 数是周期的设定，每隔多长时间执行该操作�??
         */

    }

    public void configToReceiveMail() {// throws Exception {
        try {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            ExchangeCredentials credentials = new WebCredentials("easquotation@lenovo.com", "xQaI-3962");
            service.setCredentials(credentials);
            service.setUrl(new URI("https://" + "mail.lenovo.com" + "/ews/exchange.asmx"));
            // service.autodiscoverUrl("easquotation@lenovo.com");
            Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);

            // 读取未读邮件
            SearchFilter sFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            ItemView view2 = new ItemView(Integer.MAX_VALUE);
            FindItemsResults<Item> findResults = service.findItems(WellKnownFolderName.Inbox, sFilter, view2);
            int unreadCount = 0;
            for (Item item : findResults.getItems()) {
                unreadCount++;
                System.out.println(item.getSubject());
                // 处理邮件
                EmailMessage message = EmailMessage.bind(service, item.getId());
                parseMessage(message);
                System.out.println("读取的第" + unreadCount + "封邮件主题:" + message.getSubject());
            }
            System.out.println("未读邮件数量:" + unreadCount);
            System.out.println("本次邮件读取完成，请打开Excel文件查看.");

        }
        catch (Exception exp) {
            exp.printStackTrace();
        }
    }

    public void parseMessage(EmailMessage message) throws Exception {
        try {
            // 获取发件人，发送时间，邮件内容
            System.out.println("发件人: " + getFrom(message));
            System.out.println("发送时间：" + getSentDate(message, null));
            String sentTime = getSentDate(message, null);
            String sentFrom = getFrom(message);

            // 将邮件体message的文本内容放入content?
            StringBuffer content = new StringBuffer(2000);
            getMailTextContent(message, content);

            // 调用mailSpider进行邮件的内容爬取
            MailSpider mailSpider = new MailSpider();
            mailSpider.getMailContent(content.toString(), sentTime, sentFrom);

            // 将邮件标记为已读
            message.setIsRead(true);
            message.update(ConflictResolutionMode.AlwaysOverwrite);

        }
        catch (FileNotFoundException np) {
            // 将邮件标记为未读
            message.setIsRead(false);
            message.update(ConflictResolutionMode.AlwaysOverwrite);
            System.out.println("parseMessage:Exception");
        }

    }

    /**
     * 获得邮件发件人
     * 
     * @param msg
     *            邮件内容
     * @return 姓名 <Email地址>
     * @throws MessagingException
     * @throws ServiceLocalException
     * @throws UnsupportedEncodingException
     */
    public static String getFrom(EmailMessage msg) throws MessagingException, ServiceLocalException {
        String from = "";
        String addr = "";
        try {

            EmailAddress eAddress = msg.getFrom();
            addr = eAddress.getAddress();
            String personString = eAddress.getName();
            from = personString + "<" + addr + ">";
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        return addr;
    }

    /**
     * 获得邮件发送时间
     * 
     * @param msg
     *            邮件内容
     * @return yyyy-mm-dd-HH:mm 
     * @throws MessagingException
     * @throws ServiceLocalException
     */
    public static String getSentDate(EmailMessage msg, String pattern) throws MessagingException, ServiceLocalException {
        Date receivedDate = msg.getDateTimeSent();
        if (receivedDate == null)
            return "";

        if (pattern == null || "".equals(pattern))
            pattern = "yyyy/MM/dd HH:mm ";

        return new SimpleDateFormat(pattern).format(receivedDate);
    }

    /**
     * 获得邮件文本内容
     * 
     * @param part
     *            邮件体
     * @param content
     *            存储邮件文本内容的字符串
     * @throws MessagingException
     * @throws ServiceLocalException
     * @throws IOException
     */
    public static void getMailTextContent(EmailMessage msg, StringBuffer content) throws MessagingException,
                                                                                 ServiceLocalException {
        content.append(msg.getBody().toString());
    }

    //内部类
    //计时器，是一个线程
    private class MailTask extends TimerTask {
        //每隔一定时间执行一次
        public void run() {
            configToReceiveMail();//为拉取邮件配置参数，进而开始拉取邮件

        }
    }
}
