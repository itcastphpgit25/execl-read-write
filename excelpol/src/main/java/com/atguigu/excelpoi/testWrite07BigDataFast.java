package com.atguigu.excelpoi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;

public class testWrite07BigDataFast {
    @Test
    public void testWBDF() throws Exception {
        long begin = System.currentTimeMillis();
        //创建一个SXSSFWorkbook工作簿
        Workbook workbook = new SXSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet("SXSSF测试大数据，大快少");

        //创建行
        for (int rowNum = 0; rowNum < 180000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        FileOutputStream out = new FileOutputStream("d:/excel-poi/test-write07-bigdata-fast.xlsx");

        workbook.write(out);
        out.close();

        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();

        long end = System.currentTimeMillis(); //3.887
        System.out.println("文件生成成功!!!时间共计："+(double)(end-begin)/1000);


    }
}
