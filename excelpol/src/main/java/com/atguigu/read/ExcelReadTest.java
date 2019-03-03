package com.atguigu.read;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.Date;

public class ExcelReadTest {
    @Test
    public void testRead03()throws Exception{
        //以字节方式读取文件内容
        FileInputStream is = new FileInputStream("d:/excel-poi/商品表-03.xls");
        Workbook workbook = new HSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        //读取第一行第一列
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        //输出单元内容
        System.out.println(cell.getStringCellValue());
        is.close();
    }
    @Test
    public void testRead07()throws Exception{
        FileInputStream is = new FileInputStream("d:/excel-poi/商品表-07.xlsx");
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        //读取第一行
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue());
        is.close();
    }

    //读取不同的数据类型
    @Test
    public void testCellType()throws Exception{
        FileInputStream is = new FileInputStream("d:/excel-poi/会员消费商品明细表.xls");
        //创建工作簿
        Workbook workbook = new HSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        int rowCount = sheet.getPhysicalNumberOfRows();//表格数据行数

        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            //每一行的数据
            Row rowData = sheet.getRow(rowNum);
            if(rowData!=null){
                //行的列数
                int cellCount = rowData.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("【"+(rowNum+1)+"-"+(cellNum+1)+"】");

                    //获取单元格数据
                    Cell cell = rowData.getCell(cellNum);
                    if(cell!=null){
                        int cellType = cell.getCellType();

                        //System.out.println(cellType);
                        String cellValue="";
                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING:
                                System.out.print("【字符串】");
                                //获取字符串类型值
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.print("【布尔类型】");
                                boolean booleanCellValue = cell.getBooleanCellValue();
                                cellValue = String.valueOf(booleanCellValue);
                                break;
                            case  HSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print("【数值】");
                                double numericCellValue = cell.getNumericCellValue();
                                //cellValue = String.valueOf(numericCellValue);

                                if(HSSFDateUtil.isCellDateFormatted(cell)){
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else {
                                    //或者就是手机号
                                    System.out.println("【转字符串】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                System.out.println("【数据类型错误】");
                        }
                        //输出值
                        System.out.println(cellValue);

                    }
                }
            }
        }

        is.close();



    }
}
