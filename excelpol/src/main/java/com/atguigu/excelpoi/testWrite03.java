package com.atguigu.excelpoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;

//XLS写-03
public class testWrite03 {
    @Test
    public void testWrite03() throws Exception{
        long l1 = System.currentTimeMillis();
        //创建新的Excel工作簿
        Workbook workbook = new HSSFWorkbook();
        //在工作博中创建工作表
        Sheet sheet = workbook.createSheet("会员登录统计表");

        //创建行:0代表第一行
        Row row1 = sheet.createRow(0);
        
        //创建所在行的单元格：0代表第一列
        Cell cell11 = row1.createCell(0);
        //单元格的数据
        cell11.setCellValue("今日人数");
        
        //创建单元格
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);
        
        //创建行
        Row row2 = sheet.createRow(1);

        //创建列
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        Cell cell22 = row2.createCell(1);
        String dateTime = new DateTime().toString("yyyy-MM-dd HH:mm:ss");

        cell22.setCellValue(dateTime);

        //创建---输出文件流（先创建文件夹）
        FileOutputStream outputStream = new FileOutputStream("d:/excel-poi/test-write03.xls");
        //把相应的Excel工作簿存盘
        workbook.write(outputStream);

        //关闭文件
        outputStream.close();

        long l2 = System.currentTimeMillis();
        System.out.println("文件生成成功!!!时间共计："+(l2-l1)); //261


    }
}
