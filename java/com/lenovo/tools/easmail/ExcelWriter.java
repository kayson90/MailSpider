package com.lenovo.tools.easmail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

public class ExcelWriter {
    protected String excelfile = "D:/Standalone审批模板.xls";
    protected int insertrow = 1;// 得到Excle表格写入的位置
    protected int count = 1;// 计数，审批信息（是第几条）

    protected FileInputStream in;
    protected POIFSFileSystem fs;
    protected HSSFWorkbook workbook;// 表集
    protected HSSFSheet sheet;// 工作表
    protected HSSFFont font;// 字体
    protected HSSFCellStyle style;// 格式

    public ExcelWriter() {
        // TODO Auto-generated constructor stub
    }

    @SuppressWarnings("deprecation")
    public void addContent() throws Exception {
        try {
            // 打开Excel表，读入内容到workbook,sheet
            excelfile = "D:/SoftBundle审批模板.xls";
            in = new FileInputStream(new File(excelfile));
            fs = new POIFSFileSystem(in);
            workbook = new HSSFWorkbook(fs);
            sheet = workbook.getSheetAt(0);
            font = workbook.createFont();
            style = workbook.createCellStyle();

            // int lastrow = sheet.getLastRowNum();
            // sheet.removeMergedRegion(2);
            // sheet.removeMergedRegion(37);
            // HSSFRow row=sheet.getRow(7);
            // sheet.removeRow(row);
            sheet.shiftRows(11, 12, -1);
            // short a=10;
            // sheet.removeMergedRegion(72);
            // sheet.addMergedRegion(new Region(9, a, 12, a));
            // Region d=sheet.getMergedRegionAt(72);
            // System.out .println("region70="+d.toString());
            // int ss=d.getRowFrom();
            // int xx=d.getRowTo();
            // d.setRowTo(12);
            // sheet.addMergedRegion(d);
            // System.out .println("xx= "+xx+"  ss="+ss);
            /*
             * CellRangeAddress cra = new CellRangeAddress(9, 12, 1, 1);// 合并单元格
             * sheet.addMergedRegion(cra);
             */

            // deleteRows(8, 8);
            // sheet.shiftRows(insertrow, lastrow, prolength);
            // System.out.println("sheet.getNumMergedRegions()="+sheet.getNumMergedRegions());

            // System.out.println("sheet.getPhysicalNumberOfRows()="+sheet.getPhysicalNumberOfRows());
            // System.out.println("sheet.getLastRowNum()="+sheet.getR);
            // CellRangeAddress ca = sheet.getMergedRegion(insertrow % 35 + 2);
            // int formerprolength = ca.getLastRow() - ca.getFirstRow();
            //
            // insertRow(workbook, sheet, 3, 1);
            // sheet.shiftRows(5, 8, 4);
            System.out.println("end");

            // 将内容写入Excel
            OutputStream out = new FileOutputStream(excelfile);
            out.flush();
            workbook.write(out);
            out.close();
            in.close();

        }
        catch (FileNotFoundException fp) {
            System.out.println("addContent:FileNotFoundException.");
        }
        catch (IOException iop) {
            System.out.println("addContent:IOException.");
        }
        catch (Exception p) {
            p.printStackTrace();
        }
        finally {

        }
    }

    // 写信息value到第insertrow+1行，第column1列
    public void setSingleCell(String value, int insertrow, int column1) {
        HSSFRow row;
        boolean x=(sheet.getRow(0) == null);
        if (sheet.getRow(insertrow) == null) {
            row = sheet.createRow(insertrow);
        }
        else {
            row = sheet.getRow(insertrow);
        }
        HSSFCell cell = row.createCell(column1);
        cell.setCellValue(value);
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 10);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        font.setColor(HSSFColor.BLACK.index);
        style.setFont(font);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cell.setCellStyle(style);
    }

    // 写多个单元格
    public void setMultiCell(String value, int insertrow, int column, int prolength) {
        int mergedIndex = getMergedRegionIndex(sheet, insertrow, column);
        System.out.println("mergedIndex="+mergedIndex);
        if (mergedIndex != -1) {
            sheet.removeMergedRegion(mergedIndex);
        }
        CellRangeAddress cra = new CellRangeAddress(insertrow, insertrow + prolength - 1, column, column);//CellRangeAddress(firstRow, lastRow, firstCol, lastCol)	
        //Region cra = new Region((short) insertrow, (short) column, (short) (insertrow + prolength - 1), (short) column);
        sheet.addMergedRegion(cra);
        //sheet.getMergedRegion(index);

        HSSFRow row;
        if (sheet.getRow(insertrow) == null) {
            row = sheet.createRow(insertrow);
        }
        else {
            row = sheet.getRow(insertrow);
        }
        Cell cell = row.createCell(column);
        cell.setCellValue(value);

        //short col = (short) column;
        //sheet.addMergedRegion(new Region(insertrow, col,insertrow + prolength - 1,col));//Region(rowFrom, colFrom, rowTo, colTo);

        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 10);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        font.setColor(HSSFColor.BLACK.index);
        style.setFont(font);
        style.setAlignment(HSSFCellStyle.VERTICAL_CENTER);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        setRegionBorder(HSSFCellStyle.BORDER_THIN, cra, sheet, workbook);
        cell.setCellStyle(style);

    }

    private static void setRegionBorder(int border, CellRangeAddress region, Sheet sheet, HSSFWorkbook wb) {
        RegionUtil.setBorderBottom(border, region, sheet, wb);
        RegionUtil.setBorderLeft(border, region, sheet, wb);
        RegionUtil.setBorderRight(border, region, sheet, wb);
        RegionUtil.setBorderTop(border, region, sheet, wb);

    }

    public static int getMergedRegionIndex(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress caAddress = sheet.getMergedRegion(i);
            int firstColumn = caAddress.getFirstColumn();
            int lastColumn = caAddress.getLastColumn();
            int firstRow = caAddress.getFirstRow();
            int lastRow = caAddress.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return i;
                }
            }
        }
        return -1;
    }

    // 删除sheet的第startrow到endrow行
    public void deleteRows(int startrow, int endrow) {

        for (int i = startrow; i <= endrow; i++) {
            HSSFRow deleterow = sheet.getRow(i);
            sheet.removeRow(deleterow);
        }

    }

    // 获得审批信息中记录的上一次商品/物料信息数
    /*
     * public int findFormerProLength()throws IOException{
     * 
     * 
     * return }
     */

    // 判断商机编号是否已经存在
    public boolean isExist(String number) throws IOException {

        HSSFRow indexRow = sheet.getRow(1);
        
          boolean rr=(indexRow != null); 
         
        for (int tmpRow = 1; indexRow != null;) {// 循环每一行
            HSSFCell indexCell = sheet.getRow(tmpRow).getCell(2);
            // 此行没有第2列（合并的单元格），跳过
            rr=(indexRow != null); 
            boolean cc=(indexCell != null);
            boolean op=(indexCell.getCellType() != HSSFCell.CELL_TYPE_BLANK);
            if (indexCell != null && indexCell.getCellType() != HSSFCell.CELL_TYPE_BLANK) {
                String sjbhCellValue = sheet.getRow(tmpRow).getCell(2).getStringCellValue();
                 boolean aa=number.equalsIgnoreCase(sjbhCellValue);
                if (number.equalsIgnoreCase(sjbhCellValue)) {
                    insertrow = tmpRow;
                    System.out.println("insertrow=" + insertrow);

                    HSSFCell countCell = sheet.getRow(insertrow).getCell(0);
                    count = Integer.parseInt(countCell.getStringCellValue());
                    return true;
                }
            }

            tmpRow++;
            indexRow = sheet.getRow(tmpRow);

        }
        //没有找到相同的商机编号
        insertrow = 1 + sheet.getLastRowNum();
        System.out.println("insertrow=" + insertrow);
        if (insertrow != 1) {
            int c = insertrow - 1;
            HSSFRow countRow = sheet.getRow(c);
            while ((sheet.getRow(c) == null) || (sheet.getRow(c).getCell(0) == null)||(sheet.getRow(c).getCell(0).getCellType() == HSSFCell.CELL_TYPE_BLANK)) {      	
                c--;
            }
            count = Integer.parseInt(sheet.getRow(c).getCell(0).getStringCellValue()) + 1;

        }
        System.out.println(count);
        return false;

        /*
         * } catch (FileNotFoundException fp) {
         * System.out.println("isExist:FileNotFoundException."); } catch
         * (IOException iop) { System.out.println("isExist:IOException."); }
         * catch (Exception p) { p.printStackTrace(); } finally {
         * 
         * } return false;
         */
    }

}
