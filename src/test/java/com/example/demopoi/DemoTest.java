package com.example.demopoi;

import cn.hutool.core.lang.tree.TreeUtil;
import cn.hutool.json.JSONUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.example.demopoi.listener.DemoExtraListener;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.*;

@SpringBootTest
public class DemoTest {
    @Test
    void contextLoads() {
    }

    @Test
    public void test01() {
        String fileName = "C:\\Users\\Administrator\\Desktop\\vv1.xlsx";
        EasyExcel.read(fileName, new DemoExtraListener())
                .headRowNumber(0)
                .extraRead(CellExtraTypeEnum.MERGE)
                .sheet()
                .doRead();
        Integer maxColumn = (Integer) DemoExtraListener.maxColumn.last();
        List<Map> datas = DemoExtraListener.datas;
        List<Map<String, Object>> maps1 = DemoExtraListener.maps;
        List<Map<String, Object>> allLocation = allLocation(datas.size(),maxColumn,maps1,datas);
        List<Map<String, Object>> result = buidTree(allLocation);
        String json = JSONUtil.parse(result).toString();
        System.out.println(json);
    }

    public List<Map<String, Object>> buidTree(List<Map<String, Object>> list){
        List<Map<String, Object>> tree=new ArrayList<>();
        for(Map<String, Object> node:list){
            if(Integer.parseInt(node.get("firstRowIndex").toString()) == 0){
                tree.add(findChild(node,list));
            }
        }
        return tree;
    }

    public Map<String, Object> findChild(Map<String, Object> node, List<Map<String, Object>> list){
        for(Map<String, Object> n:list){
            if(conditional(node,n)){
                if(node.get("son") == null){
                    node.put("son",new ArrayList<Map<String, Object>>());
                }
                ((List)node.get("son")).add(findChild(n,list));
            }
        }
        return node;
    }

    public boolean conditional(Map<String, Object> up,Map<String, Object> down){
        int upFirstColumnIndex = Integer.parseInt(up.get("firstColumnIndex").toString());
        int upLastColumnIndex = Integer.parseInt(up.get("lastColumnIndex").toString());
        int downFirstColumnIndex = Integer.parseInt(down.get("firstColumnIndex").toString());
        int downLastColumnIndex = Integer.parseInt(down.get("lastColumnIndex").toString());
        int upFirstRowIndex = Integer.parseInt(up.get("firstRowIndex").toString());
        int downFirstRowIndex = Integer.parseInt(down.get("firstRowIndex").toString());
        if (upFirstColumnIndex <= downFirstColumnIndex && upLastColumnIndex >= downLastColumnIndex &&
                downFirstRowIndex - upFirstRowIndex == 1) {
            return true;
        }else {
            return false;
        }
    }

    public List<Map<String,Object>> allLocation(int rowIndex,int columnIndex,List<Map<String,Object>> mergeCells,List<Map> datas){
        List<Map<String,Object>> allLocation = new ArrayList<>();
        for (int i = 0; i < rowIndex; i++) {
            for (int j = 0; j < columnIndex + 1; j++) {
                HashMap<String, Object> hashMap = new HashMap<>();
                hashMap.put("firstRowIndex", i);
                hashMap.put("lastRowIndex", i);
                hashMap.put("firstColumnIndex", j);
                hashMap.put("lastColumnIndex", j);
                for (Map<String,Object> map : mergeCells) {
                    Integer mapFirstRowIndex = Integer.parseInt(map.get("firstRowIndex").toString());
                    Integer mapLastRowIndex = Integer.parseInt(map.get("lastRowIndex").toString());
                    Integer mapFirstColumnIndex = Integer.parseInt(map.get("firstColumnIndex").toString());
                    Integer mapFastColumnIndex = Integer.parseInt(map.get("lastColumnIndex").toString());
                    Object o = datas.get(mapFirstRowIndex).get(mapFirstColumnIndex);
                    map.put("text", o);
                    if (i >= mapFirstRowIndex &&
                            i <= mapLastRowIndex &&
                            j >= mapFirstColumnIndex &&
                            j <= mapFastColumnIndex) {
                        hashMap.clear();
                    }
                }
                if (!hashMap.isEmpty()) {
                    Object o = datas.get(i).get(j);
                    hashMap.put("text", o);
                    allLocation.add(hashMap);
                }
            }
        }
        allLocation.addAll(mergeCells);
        return allLocation;
    }
}
