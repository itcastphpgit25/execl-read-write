package com.atguigu.excelpoi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;

public class testWrite07BigData {
    @Test
    public void testWBD()throws Exception{
        long begin = System.currentTimeMillis();
        //创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet("大数据插表XSSF测试");
        //创建行
        for (int rowNum = 0; rowNum < 180000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            //创建列
            for (int cellNum = 0; cellNum <10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                //单元格插入数据
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("表已插完！！！");
        //创建输出流
        FileOutputStream out = new FileOutputStream("d:/excel-poi/test-write07-bigdata.xlsx");
        //写入流
        workbook.write(out);
        out.close();
        long end = System.currentTimeMillis(); //23
        System.out.println("文件生成成功!!!时间共计："+(double)(end-begin)/1000);
    }
}
