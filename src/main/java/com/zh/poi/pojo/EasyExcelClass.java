package com.zh.poi.pojo;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
public class EasyExcelClass {

    @ExcelProperty("字符串标题")
    private String strTitle;

    @ExcelProperty("日期标题")
    private Date dateTitle;

    @ExcelProperty("数字标题")
    private Double doubleTitle;

    /**
     * 屏蔽字段
     */
    @ExcelIgnore
    private String ignore;
}
