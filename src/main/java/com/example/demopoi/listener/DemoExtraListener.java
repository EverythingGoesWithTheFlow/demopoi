package com.example.demopoi.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellExtra;

import java.util.*;

public class DemoExtraListener extends AnalysisEventListener<Map<Integer,String>> {

    public static List<Map<String,Object>> maps = new ArrayList<>();
    public static List<Map> datas = new ArrayList<>();
    public static TreeSet maxColumn = new TreeSet();
    public static Integer maxRow = 0;

    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext analysisContext) {
        datas.add(data);
        maxColumn.addAll(data.keySet());
        Integer rowIndex = analysisContext.readRowHolder().getRowIndex();
        if (maxRow == rowIndex) {
            maxRow+=1;
        }else {
            datas.remove(datas.get(maxRow));
        }
    }

    @Override
    public void extra(CellExtra extra, AnalysisContext analysisContext) {
        HashMap<String,Object> map = new HashMap();
        map.put("firstRowIndex",extra.getFirstRowIndex()+"");
        map.put("firstColumnIndex",extra.getFirstColumnIndex()+"");
        map.put("lastRowIndex",extra.getLastRowIndex()+"");
        map.put("lastColumnIndex",extra.getLastColumnIndex()+"");
        maps.add(map);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
    }

}
