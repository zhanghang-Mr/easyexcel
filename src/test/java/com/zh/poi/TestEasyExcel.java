package com.zh.poi;

import com.alibaba.excel.EasyExcel;
import com.alibaba.fastjson.JSON;
import com.zh.poi.listener.EasyExcelListener;
import com.zh.poi.pojo.EasyExcelClass;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 测试easyExcel
 */
@SpringBootTest
public class TestEasyExcel {
    //日期格式
    static SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    final static String PATH = "E:\\Demo\\shiro\\springboot-poi\\poi_demo\\";

    /**
     * easyExcel 简单的读取操作
     */
    @Test
    void testRead(){
        //设置路径
        String fileName = PATH+"easyExcel.xlsx";
        //监听器
        EasyExcelListener listener = new EasyExcelListener();
        /**
         * 这里的参数说明：
         * fileName： 读文件所在路径
         * EasyExcelClass.Class: 指定用哪个类来作为输出对象，格式类
         * listener ： 监听器，不能被String管理，每次读取都要excel重新new
         * sheet： 指定读取哪一个表，默认
         *
         */
        EasyExcel.read(fileName, EasyExcelClass.class, listener)
                .sheet()
                .doRead();

        System.out.println("---result---:"+ JSON.toJSONString(listener.getEasyExcelClassList()));
    }

    /**
     * easyexcel 写入的简单操作
     */
    @Test
    void testWrite(){
        //设置路径
        String fileName = PATH+"easyExcel.xlsx";

        /**
         *这里的参数说明：
         *  fileName： 输出的路径
         *  EasyExcelClass.Class: 指定用哪个类来作为输出对象，格式类
         *  sheetName: 模板名称
         *  doWrite： 输出的数据
         */
        EasyExcel.write(fileName, EasyExcelClass.class)
                .sheet("模板")
                .doWrite(data());
    }

    private List<EasyExcelClass> data(){
        List<EasyExcelClass> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            EasyExcelClass excel = new EasyExcelClass();
            excel.setStrTitle("字符串:"+i);
            excel.setDoubleTitle(0.56);
            excel.setDateTitle(new Date());
            excel.setIgnore("屏蔽字段："+i);
            list.add(excel);
        }
        return list;
    }
}
