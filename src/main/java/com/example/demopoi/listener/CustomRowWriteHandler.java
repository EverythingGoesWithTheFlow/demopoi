package com.example.demopoi.listener;


import com.alibaba.excel.util.WorkBookUtil;
import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;

public class CustomRowWriteHandler implements RowWriteHandler {

    @Override
    public void beforeRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Integer integer, Integer integer1, Boolean aBoolean) {
    }

    @Override
    public void afterRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer integer, Boolean aBoolean) {
        Integer integer1 = newRowIndex(integer);
        SXSSFSheet sheet = (SXSSFSheet) writeSheetHolder.getSheet();
        sheet.setRandomAccessWindowSize(100);
        row = WorkBookUtil.createRow(writeSheetHolder.getSheet(), newRowIndex(integer));
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer integer, Boolean aBoolean) {

    }

    public Integer newRowIndex(Integer integer){
        int n = integer + 5;
        return n;
    }
}
