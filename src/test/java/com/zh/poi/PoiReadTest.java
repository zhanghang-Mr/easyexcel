package com.zh.poi;


import com.alibaba.excel.EasyExcel;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 测试读数据
 */
@SpringBootTest
public class PoiReadTest {

    //日期格式
    static SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    final static String PATH = "E:\\Demo\\shiro\\springboot-poi\\poi_demo\\";

    /**
     * 读取03版 不同数据类型
     * @throws IOException
     */
    @Test
    void testCellType() throws IOException {

        //创建文件流
        FileInputStream inputStream = new FileInputStream(PATH + "poi_demo测试poi(03).xls");

        //1, 创建一个工作薄
        Workbook hssfWorkbook = new HSSFWorkbook(inputStream);

        //2，得到表, 根据下标得到表
        Sheet sheetAt = hssfWorkbook.getSheetAt(0);

        //3， 得到行, 根据下标得到行,获取标题内容
        Row rowTitle = sheetAt.getRow(0);
        if(rowTitle != null){
            //获取该行的单元格的数量
            int sellType = rowTitle.getPhysicalNumberOfCells();
            //遍历所有单元格
            for (int cellNum = 0; cellNum < sellType; cellNum++) {
                //获取当前单元格
                Cell cell = rowTitle.getCell(cellNum);
                if(cell != null){
                    //获取当前元素的数据类型
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellType+"|"+cellValue);
                }
            }
            System.out.println();
        }
        //获取表中所有行的数量
        int rowCount = sheetAt.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheetAt.getRow(rowNum);
            if(rowData != null){
                int cellCount = rowData.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("["+(rowNum+1)+"-"+(cellNum+1)+"]");

                    //获取当前元素
                    Cell cell = rowData.getCell(cellNum);
                    //匹配数据类型
                    if(cell != null){
                        //获取数据类型
                        int cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING:  //字符串
                                System.out.print("[String]");
                                cellValue = cell.getStringCellValue();
                                break;

                            case HSSFCell.CELL_TYPE_BOOLEAN: //布尔
                                System.out.print("[boolean]");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK: //为空
                                System.out.print("[blank]");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC: //数字格式
                                System.out.print("number");
                                if(HSSFDateUtil.isCellDateFormatted(cell)){
                                    System.out.print("[日期]");
                                    Date date = cell.getDateCellValue();
                                    cellValue = sf.format(date);
                                }else{
                                    // 如果不是日期格式，为了防止数字过长 转为字符串
                                    System.out.print("[转换成字符串]");
//                                    cell.setCellValue(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = String.valueOf(cell.getNumericCellValue());
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR: //错误
                                System.out.println("[error]");
                                break;
                        }
                        System.out.print(cellValue);
                        System.out.println();
                    }
                }
            }


        }
        //关闭流
        inputStream.close();
    }

    /**
     * 07 版
     * @throws IOException
     */
    @Test
    void testRead07() throws IOException {

        //创建文件流
        FileInputStream inputStream = new FileInputStream(PATH + "poi_democontextLoads07BigDataS.xlsx");

        //1, 创建一个工作薄
        Workbook hssfWorkbook = new XSSFWorkbook(inputStream);

        //2，得到表, 根据下标得到表
        Sheet sheetAt = hssfWorkbook.getSheetAt(0);

        //3， 得到行, 根据下标得到行
        Row row = sheetAt.getRow(0);

        //4， 得到列。根据下表得到列
        Cell cell = row.getCell(2);

        //注意：获取值，要判断值得类型
        //获取当前单元格的数据
        System.out.println(cell.getNumericCellValue());
        //关闭流
        inputStream.close();
    }

    /**
     * 读取03版
     */
    @Test
    void testRead03() throws IOException {

        //创建文件流
        FileInputStream inputStream = new FileInputStream(PATH + "poi_demo测试poi(03)-1.xls");

        //1, 创建一个工作薄
        Workbook hssfWorkbook = new HSSFWorkbook(inputStream);

        //2，得到表, 根据下标得到表
        Sheet sheetAt = hssfWorkbook.getSheetAt(0);

        //3， 得到行, 根据下标得到行
        Row row = sheetAt.getRow(0);

        //4， 得到列。根据下表得到列
        Cell cell = row.getCell(0);

        //注意：获取值，要判断值得类型
        //获取当前单元格的数据
        System.out.println(cell.getStringCellValue());
        //关闭流
        inputStream.close();
    }
}
