package com.lenovo.tools.easmail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.jsoup.nodes.Element;

public class SDA_SOWMail {

    private String sentTime;// 审批邮件发送时间
    private String sentFrom;// 审批邮件发送人
    private boolean completed;// 完成审批则置为true，否则为false

    private String type;// 审批类型
    private String number;// 商机编号
    private String access;// 销售通路
    private String ITcode;// 服务销售IT code
    private String client;// 签约用户名
    private String user;// 最终客户名
    private String content;// 服务内容
    private String totalmoney;// 合同总金额
    private String averagediscount;// 平均折扣
    private String GP;
    private String submitter;// 提交人

    // 审批人信息
    private class Approver {
        private String pending;// 待审批
        private String done;// 实际审批
        private String opinion;// 审批意见
    }

    Approver[] approver = new Approver[6];// 总共6个审批人

    // 构造函数，初始化属性值
    SDA_SOWMail() {
        for (int i = 0; i < 6; i++) {
            approver[i] = new Approver();
        }

    }

    // 获取table中，第row行，column列单元格的内容
    public String getCellText(Element table, int row, int column) throws Exception {
        return table.child(0).child(row).child(column).text().trim();
    }

    // 从邮件表格中读取数据到本对象中
    public void getContent(Element table, String sentTime, String sentFrom) {
        // 邮件发送时间(审批时间)，邮件发送人
        this.sentTime = sentTime;
        this.sentFrom = sentFrom;
        try {
            // 依次将表中内容读取到相应属性中
            type = getCellText(table, 0, 1);// 获得审批类型
            number = getCellText(table, 0, 3);// 获得商机编号
            access = getCellText(table, 1, 1);// 获得销售通路
            ITcode = getCellText(table, 1, 3);// 获得服务销售IT code
            client = getCellText(table, 2, 1);// 获得签约用户名
            user = getCellText(table, 3, 1);// 获得最终客户名
            content = getCellText(table, 4, 1);// 获得服务内容
            totalmoney = getCellText(table, 5, 1);// 获得合同总金额
            averagediscount = getCellText(table, 6, 1);// 获得平均折扣
            GP = getCellText(table, 6, 3);// 获得GP值
            submitter = getCellText(table, 15, 1);// 获得提交人

            // 依次获得各个审批人审批情况
            for (int n = 0; n < approver.length; n++) {
                approver[n].pending = getCellText(table, 9 + n, 2);
                approver[n].done = getCellText(table, 9 + n, 3);
                approver[n].opinion = getCellText(table, 9 + n, 4);
            }
        }
        catch (Exception p) {
            System.out.println("getContent:Exception");
        }
    }

    // 将表格内容写入excel
    public void writeToExcel() throws Exception {
        SDAExcelWriter eWriter = new SDAExcelWriter();
        eWriter.addContent();
    }

    private class SDAExcelWriter extends ExcelWriter {
        public SDAExcelWriter() {

        }

        public void addContent() throws Exception {
            //try {
            // 打开Excel表，读入内容到workbook,sheet
            excelfile = "Standalone审批模板.xls";
            in = new FileInputStream(new File(excelfile));
            fs = new POIFSFileSystem(in);
            workbook = new HSSFWorkbook(fs);
            sheet = workbook.getSheetAt(0);
            font = workbook.createFont();
            style = workbook.createCellStyle();

            // initializeHeader();

            // 是提交审批的邮件
            if (isSubmitMail()) {
                // if (true) {
                System.out.println("是提交审批的邮件");

            }
            // 是同意审批的邮件
            else if (isApprovalMail()) {
                System.out.println("是同意审批的邮件");
                // 是否审批完成
                if (isCompleted()) {
                    completed = true;// 将状态修改为完成审批
                    System.out.println("审批完成");
                }
            }

            // 插入新的内容到表格尾部
            setSingleCell(Integer.toString(count), insertrow, 0);// 写入序号
            setSingleCell(sentTime, insertrow, 1);//
            setSingleCell(number, insertrow, 2);//
            setSingleCell(ITcode, insertrow, 3);//
            setSingleCell(type, insertrow, 4);//
            setSingleCell(access, insertrow, 5);//
            setSingleCell(client, insertrow, 6);//
            setSingleCell(user, insertrow, 7);//
            setSingleCell(content, insertrow, 8);//
            setSingleCell(totalmoney, insertrow, 9);//
            setSingleCell(averagediscount, insertrow, 10);//
            setSingleCell(GP, insertrow, 11);//

            setSingleCell(approver[0].pending, insertrow, 12);
            setSingleCell(approver[0].done, insertrow, 13);
            setSingleCell(approver[0].opinion, insertrow, 14);//
            setSingleCell(approver[1].pending, insertrow, 15);
            setSingleCell(approver[1].done, insertrow, 16);
            setSingleCell(approver[1].opinion, insertrow, 17);//
            setSingleCell(approver[2].pending, insertrow, 18);
            setSingleCell(approver[2].done, insertrow, 19);
            setSingleCell(approver[2].opinion, insertrow, 20);//
            setSingleCell(approver[3].pending, insertrow, 21);
            setSingleCell(approver[3].done, insertrow, 22);
            setSingleCell(approver[3].opinion, insertrow, 23);//
            setSingleCell(approver[4].pending, insertrow, 24);
            setSingleCell(approver[4].done, insertrow, 25);
            setSingleCell(approver[4].opinion, insertrow, 26);//
            setSingleCell(approver[5].pending, insertrow, 27);
            setSingleCell(approver[5].done, insertrow, 28);
            setSingleCell(approver[5].opinion, insertrow, 29);//
            setSingleCell(submitter, insertrow, 30);//
            setSingleCell(sentTime, insertrow, 31);//
            String s = completed ? "是" : "否";
            setSingleCell(s, insertrow, 32);//

            // 将内容写入Excel
            OutputStream out = new FileOutputStream(excelfile);
            out.flush();
            workbook.write(out);
            out.close();
            in.close();

            /*} catch (FileNotFoundException fp) {
            	System.out.println("addContent:FileNotFoundException.");
            } catch (IOException iop) {
            	System.out.println("addContent:IOException.");
            } catch (Exception p) {
            	p.printStackTrace();
            } finally {

            }*/
        }

        // 是提交审批的邮件
        public boolean isSubmitMail() throws IOException {
            if ((submitter.contains(sentFrom)) && !(isExist(number))) {
                return true;
            }
            return false;
        }

        // 是进行行审批的邮件
        public boolean isApprovalMail() throws IOException {
            for (int i = 0; i < approver.length; i++) {
                if (approver[i].pending.contains(sentFrom) && isExist(number)) {

                    return true;
                }
            }
            return false;
        }

        // 审批完成
        public boolean isCompleted() throws IOException {
            for (int k = 0; k < approver.length; k++) {
                boolean zo = (!approver[k].pending.equals(""));
                boolean so = (!approver[k].done.contains("Y"));
                if ((!approver[k].pending.equals("")) && (!approver[k].done.contains("Y"))) {
                    return false;
                }
            }
            return true;

        }

        //初始化表头
        public void initializeHeader() {
            setSingleCell("序号", 0, 0);// 写入序号
            setSingleCell("发邮件时间", 0, 1);// 
            setSingleCell("商机编号", 0, 2);// 
            setSingleCell("服务销售IT code", 0, 3);// 
            setSingleCell("审批类型", 0, 4);// 
            setSingleCell("销售通路", 0, 5);// 
            setSingleCell("签约客户名", 0, 6);// 
            setSingleCell("最终用户名称", 0, 7);// 
            setSingleCell("服务内容", 0, 8);// 
            setSingleCell("平均折扣", 0, 9);// 
            setSingleCell("合同总金额", 0, 10);// 
            setSingleCell("GP", 0, 11);// 
            setSingleCell("一级需审批", 0, 12);// 
            setSingleCell("一级实际审批", 0, 13);// 
            setSingleCell("一级审批意见", 0, 14);// 
            setSingleCell("二级需审批", 0, 15);// 
            setSingleCell("二级实际审批", 0, 16);// 
            setSingleCell("二级审批意见", 0, 17);// 
            setSingleCell("三级需审批1", 0, 18);// 
            setSingleCell("三级实际审批1", 0, 19);// 
            setSingleCell("三级审批意见1", 0, 20);// 
            setSingleCell("三级需审批2", 0, 21);// 
            setSingleCell("三级实际审批2", 0, 22);// 
            setSingleCell("三级审批意见2", 0, 23);// 
            setSingleCell("四级需审批1", 0, 24);// 
            setSingleCell("四级实际审批1", 0, 25);// 
            setSingleCell("四级审批意见1", 0, 26);// 
            setSingleCell("四级需审批2", 0, 27);// 
            setSingleCell("四级实际审批2", 0, 28);// 
            setSingleCell("四级审批意见2", 0, 29);// 
            setSingleCell("提交人", 0, 30);// 
            setSingleCell("审批时间", 0, 31);// 
            setSingleCell("是否完成审批", 0, 32);// 

        }

    }

}
