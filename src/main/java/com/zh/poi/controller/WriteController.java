package com.zh.poi.controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteController {

    final static String PAHT = "E:\\Demo\\shiro\\springboot-poi\\poi_demo";
    public void test() throws IOException {
        //1， 创建工作薄(03)版
        Workbook workbook = new HSSFWorkbook();
        //2, 创建一个工作表
        Sheet sheet = workbook.createSheet("测试poi(03)版");
        //3, 创建一个行(1,1)
        Row row = sheet.createRow(0);
        //4, 创建一个单元格
        Cell cell = row.createCell(0);
        cell.setCellValue("(1,1)");
        Cell cell1 = row.createCell(1);
        cell1.setCellValue("(1,2)");

        // 第二行
        Row row2 = sheet.createRow(1);
        Cell cell2 = row2.createCell(0);
        Cell cell3 = row2.createCell(1);
        cell2.setCellValue("(2,1)");
        cell2.setCellValue("(2,2)");

        // 生成一张表（io流）
        FileOutputStream fileOutputStream = new FileOutputStream(PAHT + "测试poi(03).xls");
        // 输出
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();
        System.out.println("excel表生成完毕");
    }
}
