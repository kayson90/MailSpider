package com.lenovo.tools.easmail;

/**
 * Lenovo Group
 * Copyright (c) 1999-2016 All Rights Reserved.
 */

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/** 
 * 
 * @author kayson
 * @version $Id: MailSpider.java, v 0.1 2016�?3�?29�? 下午5:56:46 kayson Exp $
 */
public class MailSpider {

    /* private HashMap<String, String> mailSoftBundle = new HashMap<String, String>();
     private HashMap<String, String> mailHardBundle = new HashMap<String, String>();

     private static final Logger log = Logger.getLogger(MailSpider.class);*/

    //输入文件
    public void getMailContent(String html, String sentTime, String sentFrom) throws Exception {
        //加载文档html成为一个Document
        Document doc = Jsoup.parse(html, "GB2312");
        //获取所有的表格
        Elements tables = doc.select("table");
        int i = 0;
        for (Element table : tables) {//对于每一个表格
            //if(table.)
            if (!getCellText(table, 0, 0).contains("审批类型")) {
                continue;
            }
            //获取表的类型
            String orderType = getCellText(table, 0, 1);
            System.out.println("~~~~~~~~~~~~~~表格" + i + "类型:" + orderType);

            //是SOW审批
            if (orderType.equalsIgnoreCase("SOW审批") || (orderType.equalsIgnoreCase("SDA价格审批"))) {
                System.out.println("是SOW/SDA审批表" + orderType);
                SDA_SOWMail sow = new SDA_SOWMail();
                sow.getContent(table, sentTime, sentFrom);
                sow.writeToExcel();
            }

            //是SB价格审批
            else if (orderType.equalsIgnoreCase("SB价格审批")) {
                System.out.println("是SB审批表" + orderType);
                SBMail sb = new SBMail();
                sb.getContent(table, sentTime, sentFrom);
                sb.writeToExcel();
            }
            else {
                continue;
            }
            i++;
        }//end of for(每个表格）

    }

    //获取table中，第row行，column列单元格的内容
    public String getCellText(Element table, int row, int column) {
        return table.child(0).child(row).child(column).text().trim();
    }

}
