package com.zh.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

@SpringBootTest
class PoiDemoApplicationTests {
	//日期格式
	static SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	final static String PATH = "E:\\Demo\\shiro\\springboot-poi\\poi_demo";


	/**
	 * SXSSF 属于07版的升级版，会产生临时文件优化速度
	 * 10万条数据消耗的时间为：1.717s
	 * @throws IOException
	 */
	@Test
	void contextLoads07BigDataS() throws IOException {
		//获取当前系统时间
		long start = System.currentTimeMillis();

		//1， 创建工作簿
		Workbook workbook = new SXSSFWorkbook();

		//2, 创建一个表
		Sheet sheet = workbook.createSheet();
		//写入数据,o3版最大65536行数据
		for (int rowsNumber = 0; rowsNumber < 100000; rowsNumber++) {
			//3, 创建行
			Row row = sheet.createRow(rowsNumber);
			for (int cellNumber = 0; cellNumber < 10; cellNumber++) {
				//4， 创建列
				Cell cell = row.createCell(cellNumber);
				cell.setCellValue(cellNumber);
			}
		}

		//创建流
		FileOutputStream outputStream = new FileOutputStream(PATH + "contextLoads07BigDataS.xlsx");

		//输出流
		workbook.write(outputStream);

		//清除临时文件
		((SXSSFWorkbook)workbook).dispose();

		//关闭流
		outputStream.close();

		//结束时间
		long end = System.currentTimeMillis();

		System.out.println("--07消耗时间---："+(double)(end-start)/1000);
	}

	/**
	 * 07 版写入 大文本写入数据消耗的时间
	 * 当写入的数据行为65536 时，所消耗的时间为6.318s,比03版要慢
	 * 但是07版不限制行的数量，可以超过03版的最大行数
	 * 比如写入10万行，所消耗的时间为：12.152s
	 * @throws IOException
	 */
	@Test
	void contextLoads07BigDate() throws IOException {
		//获取当前系统时间
		long start = System.currentTimeMillis();

		//1， 创建工作簿
		Workbook workbook = new XSSFWorkbook();

		//2, 创建一个表
		Sheet sheet = workbook.createSheet();
		//写入数据,o3版最大65536行数据
		for (int rowsNumber = 0; rowsNumber < 100000; rowsNumber++) {
			//3, 创建行
			Row row = sheet.createRow(rowsNumber);
			for (int cellNumber = 0; cellNumber < 10; cellNumber++) {
				//4， 创建列
				Cell cell = row.createCell(cellNumber);
				cell.setCellValue(cellNumber);
			}
		}

		//创建流
		FileOutputStream outputStream = new FileOutputStream(PATH + "contextLoads07BigDate1.xlsx");

		//输出流
		workbook.write(outputStream);

		//关闭流
		outputStream.close();

		//结束时间
		long end = System.currentTimeMillis();

		System.out.println("--07消耗时间---："+(double)(end-start)/1000);
	}

	/**
	 * 测试03版 大文本写入数据消耗的时间
	 * 结果：消耗了1.008s
	 * 注意：如果 创建的行大于65536，
	 * 会报java.lang.IllegalArgumentException: Invalid row number (65536) outside allowable range (0..65535)错误
	 */
	@Test
	void contextLoads03BigDate() throws IOException {
		//获取当前系统时间
		long start = System.currentTimeMillis();

		//1， 创建工作簿
		HSSFWorkbook workbook = new HSSFWorkbook();

		//2, 创建一个表
		HSSFSheet sheet = workbook.createSheet();
		//写入数据,o3版最大65536行数据
		for (int rowsNumber = 0; rowsNumber < 65536; rowsNumber++) {
			//3, 创建行
			HSSFRow row = sheet.createRow(rowsNumber);
			for (int cellNumber = 0; cellNumber < 10; cellNumber++) {
				//4， 创建列
				HSSFCell cell = row.createCell(cellNumber);
				cell.setCellValue(cellNumber);
			}
		}

		//创建流
		FileOutputStream outputStream = new FileOutputStream(PATH + "contextLoads03BigDate1.xls");

		//输出流
		workbook.write(outputStream);

		//关闭流
		outputStream.close();

		//结束时间
		long end = System.currentTimeMillis();

		System.out.println("--03消耗时间---："+(double)(end-start)/1000);
	}

	/**
	 * 07版 excel 写入数据
	 * @throws IOException
	 */
	@Test
	void contextLoads07() throws IOException {

		//1， 创建工作薄(07)版
		Workbook workbook = new XSSFWorkbook();
		//2, 创建一个工作表
		Sheet sheet = workbook.createSheet("测试poi(07)版");
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
		cell3.setCellValue(sf.format(new Date()));

		// 生成一张表（io流）
		FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试poi(07)-1.xlsx");
		// 输出
		workbook.write(fileOutputStream);
		// 关闭流
		fileOutputStream.close();
		System.out.println("excel表生成完毕");
	}


	/**
	 * 03 版excel 写入数据
	 * @throws IOException
	 */
	@Test
	void contextLoads() throws IOException {

			//1， 创建工作薄(03)版
			Workbook workbook = new HSSFWorkbook();
			//2, 创建一个工作表
			Sheet sheet = workbook.createSheet("测试poi(03)版");
			//3, 创建一个行(1,1)
			Row row = sheet.createRow(0);
			//4, 创建一个单元格
			Cell cell = row.createCell(0);
			cell.setCellValue("名字");
			Cell cell1 = row.createCell(1);
			cell1.setCellValue("年龄");
			Cell cell2 = row.createCell(2);
			cell2.setCellValue("日期");

			// 第二行
			Row row2 = sheet.createRow(1);
			Cell cell0 = row2.createCell(0);
			Cell cell22 = row2.createCell(1);
			Cell cell33 = row2.createCell(2);
			cell0.setCellValue("张三");
			cell22.setCellValue(18);
			cell33.setCellValue(sf.format(new Date()));
//			cell3.setCellValue(sf.format(new Date()));

			// 生成一张表（io流）
			FileOutputStream fileOutputStream = new FileOutputStream(PATH + "测试poi(03).xls");
			// 输出
			workbook.write(fileOutputStream);
			// 关闭流
			fileOutputStream.close();
			System.out.println("excel表生成完毕");
	}

}
