package com.example.demopoi;

import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.alibaba.excel.EasyExcel;
import com.example.demopoi.listener.CustomRowWriteHandler;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.*;

@SpringBootTest
class DemopoiApplicationTests {

    @Test
    void contextLoads() {
    }

    @Test
    public void test01(){
//        String fileName = "C:\\Users\\Administrator\\Desktop\\vv1.xlsx";
//        EasyExcel.read(fileName, new DemoExtraListener())
//                .headRowNumber(0)
//                .extraRead(CellExtraTypeEnum.MERGE)
//                .ignoreEmptyRow(false)
//                .sheet().doRead();
//        Integer o = (Integer) DemoExtraListener.maxColumn.last();
//        List<Map<String,String>> allLocation = new ArrayList<>();
//        for (int i = 0; i < DemoExtraListener.datas.size(); i++) {
//            for (int j = 0; j < o+1; j++) {
//                HashMap<String, String> hashMap = new HashMap<>();
//                hashMap.put("firstRowIndex", i+"");
//                hashMap.put("lastRowIndex", i+"");
//                hashMap.put("firstColumnIndex", j+"");
//                hashMap.put("lastColumnIndex", j+"");
//                allLocation.add(hashMap);
//            }
//        }
//        Set<Map<String,String>> newMaps = new HashSet<>();
//        newMaps.addAll(DemoExtraListener.maps);
//        newMaps.addAll(allLocation);
//        for (Map<String,String> map : DemoExtraListener.maps) {
//            for (Map<String,String> al : allLocation) {
//                Integer alFirstRowIndex = Integer.parseInt(al.get("firstRowIndex"));
//                Integer alLastRowIndex = Integer.parseInt(al.get("lastRowIndex"));
//                Integer alFirstColumnIndex = Integer.parseInt(al.get("firstColumnIndex"));
//                Integer alFastColumnIndex = Integer.parseInt(al.get("lastColumnIndex"));
//                Integer mapFirstRowIndex = Integer.parseInt(map.get("firstRowIndex"));
//                Integer mapLastRowIndex = Integer.parseInt(map.get("lastRowIndex"));
//                Integer mapFirstColumnIndex = Integer.parseInt(map.get("firstColumnIndex"));
//                Integer mapFastColumnIndex = Integer.parseInt(map.get("lastColumnIndex"));
//                if (alFirstRowIndex >= mapFirstRowIndex &&
//                        alLastRowIndex <= mapLastRowIndex &&
//                        alFirstColumnIndex >= mapFirstColumnIndex &&
//                        alFastColumnIndex <= mapFastColumnIndex) {
//                    newMaps.remove(al);
//                }
//            }
//        }
//        List<List<Map<String,String>>> result = new ArrayList<>();
//        TreeSet<Map<String, String>> maps = new TreeSet<>((o1, o2) -> {
//            int firstRowIndexOne = Integer.parseInt(o1.get("firstRowIndex"));
//            int firstRowIndexTwo = Integer.parseInt(o2.get("firstRowIndex"));
//            int firstColumnIndexOne = Integer.parseInt(o1.get("firstColumnIndex"));
//            int firstColumnIndexTwo = Integer.parseInt(o2.get("firstColumnIndex"));
//            int ret = 0;
//            int sg = firstRowIndexOne - firstRowIndexTwo;
//            if (sg != 0) {
//                ret = sg > 0 ? 1 : -1;
//            } else {
//                sg = (firstColumnIndexOne - firstColumnIndexTwo) > 0 ? 1 : -1;
//                if (sg != 0) {
//                    ret = sg > 0 ? 1 : -1;
//                }
//            }
//            return ret;
//        });
//        maps.addAll(newMaps);
//        for (int i = 0; i < DemoExtraListener.datas.size(); i++) {
//            ArrayList<Map<String,String>> objects = new ArrayList<>();
//            for (Map<String, String> map : maps) {
//                if (map.get("firstRowIndex").equals(""+i)) {
//                    Object firstColumnIndex = DemoExtraListener.datas.get(i).get(Integer.parseInt(map.get("firstColumnIndex")));
//                    String s = firstColumnIndex == null ? "" : firstColumnIndex.toString();
//                    map.put("text",s);
//                    objects.add(map);
//                }
//            }
//            result.add(i,objects);
//        }
//        String json = JSONUtil.parse(result).toString();
//        System.out.println(json);


        //写文件
        FileUtil.copy("C:\\Users\\Administrator\\Desktop\\vv1.xlsx","C:\\Users\\Administrator\\Desktop\\vv2.xlsx" , true);
        //通过工具类创建writer
        List<List<Object>> lists = dataList();
        ExcelWriter writer = ExcelUtil.getWriter("C:\\Users\\Administrator\\Desktop\\vv2.xlsx");
        Sheet sheet = writer.getSheet();
        int lastRowNum = sheet.getLastRowNum();
        sheet.shiftRows(4, lastRowNum, lists.size());
        //跳过当前行，既第一行，非必须，在此演示用
        writer.passRows(4);
        //一次性写出内容，强制输出标题
        writer.write(lists, false);
        //关闭writer，释放内存
        writer.close();
    }

    private List<List<Object>> dataList() {
        List<List<Object>> list = new ArrayList<List<Object>>();
        for (int i = 0; i < 10; i++) {
            List<Object> data = new ArrayList<Object>();
            data.add("字符串" + i);
            data.add("addsadsa");
            data.add(0.56);
            list.add(data);
        }
        return list;
    }

}
