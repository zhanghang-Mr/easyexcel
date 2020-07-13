package com.zh.poi.listener;


import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import com.zh.poi.pojo.EasyExcelClass;

import java.util.ArrayList;
import java.util.List;

/**
 * easyExcel 读数据的监听器
 */
public class EasyExcelListener extends AnalysisEventListener<EasyExcelClass> {

    List<EasyExcelClass> list = new ArrayList<EasyExcelClass>();

    public  List<EasyExcelClass> getEasyExcelClassList(){
        return this.list;
    }

    /**
     * 读取数据会执行 invoke 方法
     * @param easyExcelClass  类型
     * @param analysisContext  分析上文
     */
    @Override
    public void invoke(EasyExcelClass easyExcelClass, AnalysisContext analysisContext) {
        System.out.println("--data--:"+ JSON.toJSONString(easyExcelClass));
        list.add(easyExcelClass);
    }

    /**
     * 所有的数据解析完成会调用doAfterAllAnalysed
     * @param analysisContext
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("--list--:"+ JSON.toJSONString(list));
    }
}
