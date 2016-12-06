package com.lenovo.tools.easmail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.LinkedList;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.Region;
import org.jsoup.nodes.Element;

public class SBMail {
    private String sentTime;// 审批邮件发送时间
    private String sentFrom;// 审批邮件发送人
    private boolean completed;// 完成审批则置为true，否则为false

    private String type;// 审批类型
    private String number;// 商机编号
    private String access;// 销售通路
    private String ITcode;// 服务销售IT code
    private String agency;// 分销/代理名称
    private String user;// 最终用户名

    private String totalmoney;// 合同总金额
    private String averagediscount;// 平均折扣
    private String serverGP;// 服务GP
    private String hardwarediscount;// 硬件平均折扣
    private String hardwareGP;// 硬件GP
    private String submitter;// 提交人
    private String maintainnumber;// 商务维护订单号

    // 审批人信息
    private class Approver {
        private String pending;// 待审批
        private String done;// 实际审批
        private String opinion;// 审批意见
    }

    Approver[] approver = new Approver[6];// 总共6个审批人

    // 产品/服务信息
    private class Product {
        private String namenumber;// 产品名称和物料编号
        private String quantity;// 服务数量
        private String marketprice;// 销售经理单价
        private String discountprice;// 折扣单价
        private String discount;// 折扣率
    }

    private LinkedList<Product> productslist = new LinkedList<Product>();// 使用链表存储产品信息
    private int prolength = 0;// 物料/产品信息数目
    private int formerprolength = 0;// 上一次的物料/产品信息数目，用于更新excel表格

    SBMail() {
        for (int i = 0; i < 6; i++) {
            approver[i] = new Approver();
        }
    }

    // 获取table中，第row行，column列单元格的内容
    public String getCellText(Element table, int row, int column) {
        return table.child(0).child(row).child(column).text().trim();
    }

    // 从邮件表格中读取数据到本对象中
    public void getContent(Element table, String sentTime, String sentFrom) {
        try {
            // 邮件发送时间(审批时间)，邮件发送人
            this.sentTime = sentTime;
            this.sentFrom = sentFrom;

            // 依次将表中内容读取到相应属性中
            type = getCellText(table, 0, 1);// 获得审批类型
            number = getCellText(table, 0, 3);// 获得商机编号
            access = getCellText(table, 1, 1);// 获得销售通路
            ITcode = getCellText(table, 1, 3);// 获得服务销售IT code;// 服务销售IT code
            agency = getCellText(table, 2, 1);// 获得分销/代理名称
            user = getCellText(table, 3, 1);// 获得最终客户名

            // 依次获得各产品信息项
            // formerprolength = prolength;// 更新物料/产品信息数
            prolength = 0;
            String s = getCellText(table, prolength + 5, 0);
            for (; !(s.contains("合同总金额")); prolength++) {
                // 创建新的产品对象，存入产品信息
                Product product = new Product();
                product.namenumber = s;// 产品名称和物料编号
                product.quantity = getCellText(table, prolength + 5, 1);// 服务数量
                product.marketprice = getCellText(table, prolength + 5, 2);// 销售经理单价
                product.discountprice = getCellText(table, prolength + 5, 3);// 折扣单价
                product.discount = getCellText(table, prolength + 5, 4);// 折扣率
                // 将产品放入表单信息的产品链表中
                productslist.add(product);
                s = getCellText(table, prolength + 1 + 5, 0);
            }

            totalmoney = getCellText(table, prolength + 5, 1);// 获得合同总金额9
            averagediscount = getCellText(table, prolength + 6, 2);// 获得平均折扣10
            serverGP = getCellText(table, prolength + 6, 4);// 获得服务GP
            hardwarediscount = getCellText(table, prolength + 7, 2);// 获得硬件平均折扣11
            hardwareGP = getCellText(table, prolength + 7, 4);// 获得硬件GP

            // 依次获得各个审批人审批情况
            int asslength = approver.length;// 审批人数

            for (int k = 0; k < asslength; k++) {
                approver[k].pending = getCellText(table, k + prolength + 10, 2);
                approver[k].done = getCellText(table, k + prolength + 10, 3);
                approver[k].opinion = getCellText(table, k + prolength + 10, 4);
            }

            submitter = getCellText(table, prolength + asslength + 10, 1);// 获得提交人20
            maintainnumber = getCellText(table, prolength + asslength + 10, 4);// 获得商务维护订单号
        }
        catch (Exception p) {
            System.out.println("getContent:Exception");
        }
    }

    // 将表格内容写入excel
    public void writeToExcel() throws Exception {
        // try {

        SBExcelWriter eWriter = new SBExcelWriter();
        eWriter.addContent();

        /*
         * } catch (Exception p) { System.out.println("writeToExcel:Exception");
         * }
         */

    }

    private class SBExcelWriter extends ExcelWriter {
        public SBExcelWriter() {
        }

        public void addContent() throws Exception {
            // try {
            // 打开Excel表，读入内容到workbook,sheet
            excelfile = "SoftBundle审批模板.xls";
            System.out.println(excelfile);
            in = new FileInputStream(new File(excelfile));
            fs = new POIFSFileSystem(in);
            workbook = new HSSFWorkbook(fs);
            sheet = workbook.getSheetAt(0);
            font = workbook.createFont();
            style = workbook.createCellStyle();

            // initializeHeader();

            // 是提交审批的邮件
            if (isSubmitMail()) {
                System.out.println("是提交审批的邮件");
            }
            // 是同意审批的邮件
            else if (isApprovalMail()) {
                System.out.println("是同意审批的邮件");

                // 得到原来的产品/物料数
                /*
                 * CellRangeAddress a = sheet.getMergedRegion((insertrow-1)*
                 * 35); int ss = a.getFirstRow(); int xx = a.getLastRow(); int
                 * formerprolength = a.getLastRow() - a.getFirstRow();
                 */

                Region a = sheet.getMergedRegionAt((count - 1) * 35);
                int ss = a.getRowFrom();
                int xx = a.getRowTo();
                int formerprolength = xx - ss + 1;
                System.out.println("formerprolength =" + formerprolength + "   prolength=" + prolength);

                // 清除原来的记录
                if (prolength < formerprolength) {
                    deleteRows(insertrow, insertrow + formerprolength - 1);
                    sheet.shiftRows(insertrow + formerprolength, sheet.getLastRowNum(), prolength - formerprolength);
                    System.out.println("sheet.shiftRows(" + (insertrow + formerprolength) + ","
                                       + (sheet.getLastRowNum()) + "," + (prolength - formerprolength) + ")");
                }
                // 物料增加
                else if (prolength > formerprolength && ((insertrow + formerprolength) <= sheet.getLastRowNum())) {
                    sheet.shiftRows(insertrow + formerprolength, sheet.getLastRowNum(), prolength - formerprolength);
                    System.out.println("sheet.shiftRows(" + (insertrow + formerprolength) + ","
                                       + (sheet.getLastRowNum()) + "," + (prolength - formerprolength) + ")");
                }

                // 是否审批完成
                if (isCompleted()) {
                    completed = true;// 将状态修改为完成审批
                    System.out.println("审批完成");
                }
            }

            // 插入新的内容到表格尾部
            System.out.println("insertrow=" + insertrow);
            setMultiCell(Integer.toString(count), insertrow, 0, prolength);// 写入序号
            setMultiCell(sentTime, insertrow, 1, prolength);//发送时间
            setMultiCell(number, insertrow, 2, prolength);//商机编号
            setMultiCell(ITcode, insertrow, 3, prolength);//
            setMultiCell(type, insertrow, 4, prolength);//审批类型
            setMultiCell(access, insertrow, 5, prolength);//销售通路
            setMultiCell(agency, insertrow, 6, prolength);//分销/代理
            setMultiCell(user, insertrow, 7, prolength);//最终客户名

            // 写入物料/产品信息
            for (int n = 0; n < prolength; n++) {
                Product p = productslist.get(n);
                setSingleCell(p.namenumber, insertrow + n, 8);
                setSingleCell(p.quantity, insertrow + n, 9);
                setSingleCell(p.marketprice, insertrow + n, 10);
                setSingleCell(p.discountprice, insertrow + n, 11);
                setSingleCell(p.discount, insertrow + n, 12);
            }

            setMultiCell(totalmoney, insertrow, 13, prolength);//
            setMultiCell(averagediscount, insertrow, 14, prolength);//
            setMultiCell(serverGP, insertrow, 15, prolength);//
            setMultiCell(hardwarediscount, insertrow, 16, prolength);//
            setMultiCell(hardwareGP, insertrow, 17, prolength);//

            setMultiCell(approver[0].pending, insertrow, 18, prolength);
            setMultiCell(approver[0].done, insertrow, 19, prolength);
            setMultiCell(approver[0].opinion, insertrow, 20, prolength);//
            setMultiCell(approver[1].pending, insertrow, 21, prolength);
            setMultiCell(approver[1].done, insertrow, 22, prolength);
            setMultiCell(approver[1].opinion, insertrow, 23, prolength);//
            setMultiCell(approver[4].pending, insertrow, 24, prolength);
            setMultiCell(approver[4].done, insertrow, 25, prolength);
            setMultiCell(approver[4].opinion, insertrow, 26, prolength);//
            setMultiCell(approver[5].pending, insertrow, 27, prolength);
            setMultiCell(approver[5].done, insertrow, 28, prolength);
            setMultiCell(approver[5].opinion, insertrow, 29, prolength);//
            setMultiCell(approver[2].pending, insertrow, 30, prolength);
            setMultiCell(approver[2].done, insertrow, 31, prolength);
            setMultiCell(approver[2].opinion, insertrow, 32, prolength);//
            setMultiCell(approver[3].pending, insertrow, 33, prolength);
            setMultiCell(approver[3].done, insertrow, 34, prolength);
            setMultiCell(approver[3].opinion, insertrow, 35, prolength);//
            setMultiCell(submitter, insertrow, 36, prolength);//
            setMultiCell(maintainnumber, insertrow, 37, prolength);//
            setMultiCell(sentTime, insertrow, 38, prolength);//
            String s = completed ? "是" : "否";
            setMultiCell(s, insertrow, 39, prolength);//

            // 将内容写入Excel
            OutputStream out = new FileOutputStream(excelfile);
            out.flush();
            workbook.write(out);
            out.close();
            in.close();
            /*
             * } catch (FileNotFoundException fp) {
             * System.out.println("addContent:FileNotFoundException."); } catch
             * (IOException iop) {
             * System.out.println("addContent:IOException."); } catch (Exception
             * p) { p.printStackTrace(); }
             */
        }

        // 是提交审批的邮件
        public boolean isSubmitMail() throws IOException {

            //if ((submitter.contains(sentFrom)) && !(isExist(number))) {
            if (!isExist(number)) {
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
            setSingleCell("分销/代理名称", 0, 6);// 
            setSingleCell("最终客户名", 0, 7);// 

            setSingleCell("产品名称和物料编号", 0, 8);// 
            setSingleCell("服务数量", 0, 9);// 
            setSingleCell("销售经理单价", 0, 10);// 
            setSingleCell("折扣单价", 0, 11);// 
            setSingleCell("折扣率", 0, 12);// 

            setSingleCell("合同总金额", 0, 13);// 
            setSingleCell("平均折扣", 0, 14);// 
            setSingleCell("服务GP", 0, 15);// 
            setSingleCell("硬件平均折扣", 0, 16);// 
            setSingleCell("硬件GP", 0, 17);// 

            setSingleCell("一级需审批", 0, 18);// 
            setSingleCell("一级实际审批", 0, 19);// 
            setSingleCell("一级审批意见", 0, 20);// 
            setSingleCell("二级需审批", 0, 21);// 
            setSingleCell("二级实际审批", 0, 22);// 
            setSingleCell("二级审批意见", 0, 23);// 
            setSingleCell("三级需审批1", 0, 24);// 
            setSingleCell("三级实际审批1", 0, 25);// 
            setSingleCell("三级审批意见1", 0, 26);// 
            setSingleCell("三级需审批2", 0, 27);// 
            setSingleCell("三级实际审批2", 0, 28);// 
            setSingleCell("三级审批意见2", 0, 29);// 
            setSingleCell("四级需审批1", 0, 30);// 
            setSingleCell("四级实际审批1", 0, 31);// 
            setSingleCell("四级审批意见1", 0, 32);// 
            setSingleCell("四级需审批2", 0, 33);// 
            setSingleCell("四级实际审批2", 0, 34);// 
            setSingleCell("四级审批意见2", 0, 35);// 

            setSingleCell("提交人", 0, 36);// 商务维护订单号
            setSingleCell("商务维护订单号", 0, 37);//
            setSingleCell("审批时间", 0, 38);// 
            setSingleCell("是否完成审批", 0, 39);// 
            HSSFRow head = sheet.getRow(0);

        }

    }

}
