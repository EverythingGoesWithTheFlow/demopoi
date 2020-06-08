package com.example.demopoi.utils;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class POIReadExcelToHtmlUtil {
    /**
     * excel转html入口
     * @param filePath
     * @param isWithStyle
     * @return
     */
    public static List<Map<String, String>> readExcelToHtml(String filePath, boolean isWithStyle) {
        List<Map<String, String>> excelInfoMapList = null;
        // 文件对象
        File file = new File(filePath);
        // 文件流
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(file);
            // 创建工作簿
            Workbook workbook = WorkbookFactory.create(inputStream);
            // Excel类型
            if (workbook instanceof HSSFWorkbook) {
                // 2003
                HSSFWorkbook hssfWorkbook = (HSSFWorkbook) workbook;
                // 获取Excel信息
                excelInfoMapList = getExcelInfo(hssfWorkbook, isWithStyle);
            } else if (workbook instanceof XSSFWorkbook) {
                // 2007
                XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
                // 获取Excel信息
                excelInfoMapList = getExcelInfo(xssfWorkbook, isWithStyle);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return excelInfoMapList;
    }

    /**
     * 获取Excel信息
     * @param workbook
     * @param isWithStyle
     * @return
     */
    private static List<Map<String, String>> getExcelInfo(Workbook workbook, boolean isWithStyle) {
        List<Map<String, String>> htmlMapList = new ArrayList<Map<String, String>>();
        // 获取所有sheet
        int sheets = workbook.getNumberOfSheets();
        // 遍历sheets
        for (int sheetIndex = 0; sheetIndex < sheets; sheetIndex++) {
            // 用于保存sheet信息
            Map<String, String> sheetMap = new HashMap<String, String>();
            // 获取sheet名
            String sheetName = workbook.getSheetName(sheetIndex);
            // 存储sheet名
            sheetMap.put("sheetName", sheetName);
            StringBuffer stringBuffer = new StringBuffer();
            // 获取第一个sheet信息
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            // 行数
            int lastRowNum = sheet.getLastRowNum();
            // 获取合并后的单元格行列坐标
            Map<String, String> map[] = getRowSpanColSpan(sheet);
            stringBuffer.append("<table style='border-collapse:collapse;' width='100%'>");
            Row row = null;
            Cell cell = null;
            for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
                // 获取行
                row = sheet.getRow(rowNum);
                // 空行
                if (row == null) {
                    stringBuffer.append("<tr><td> </td></tr>");
                    continue;
                }
                stringBuffer.append("<tr>");
                // 列数
                short lastCellNum = row.getLastCellNum();
                for (int colNum = 0; colNum <= lastCellNum; colNum++) {
                    // 获取列
                    cell = row.getCell(colNum);
                    // 空白单元格
                    if (cell == null) {
                        stringBuffer.append("<td> </td>");
                        continue;
                    }
                    // 获取列值
                    String cellValue = getCellValue(cell);
                    if (map[0].containsKey(rowNum + "," + colNum)) {
                        String point = map[0].get(rowNum + "," + colNum);
                        map[0].remove(rowNum + "," + colNum);
                        int bottomRow = Integer.valueOf(point.split(",")[0]);
                        int bottomCol = Integer.valueOf(point.split(",")[1]);
                        int rowSpan = bottomRow - rowNum + 1;
                        int colSpan = bottomCol - colNum + 1;
                        stringBuffer.append("<td rowspan= '" + rowSpan + "' colSpan= '" + colSpan + "' ");
                    } else if (map[1].containsKey(rowNum + "," + colNum)) {
                        map[1].remove(rowNum + "," + colNum);
                        continue;
                    } else {
                        stringBuffer.append("<td ");
                    }

                    // 判断是否包含样式
                    if (isWithStyle) {
                        // 处理单元格样式
                        dealExcelStyle(workbook, sheet, cell, stringBuffer);
                    }

                    stringBuffer.append(">");
                    if (cellValue == null || "".equals(cellValue.trim())) {
                        stringBuffer.append("   ");
                    } else {
                        stringBuffer.append(cellValue.replace(String.valueOf((char) 160), " "));
                    }
                    stringBuffer.append("</td>");
                }
                stringBuffer.append("</tr>");
                if (rowNum > 500) {
                    stringBuffer.append("<tr><td>数据量太大，请下载Excel查看更多数据</td></tr>");
                    break;
                }
            }
            stringBuffer.append("</table>");
            sheetMap.put("content", stringBuffer.toString());
            htmlMapList.add(sheetMap);
        }
        return htmlMapList;
    }

    /**
     * 获取列值
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell) {
        String result = new String();
        if (cell.getCellTypeEnum() == CellType.NUMERIC) {   // 数字类型
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat simpleDateFormat = null;
                // 时间
                if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                    simpleDateFormat = new SimpleDateFormat("HH:mm");
                } else {
                    // 日期
                    simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
                }
                Date date = cell.getDateCellValue();
                result = simpleDateFormat.format(date);
            } else if (cell.getCellStyle().getDataFormat() == 58) {
                // 处理自定义日期格式：m月d日（通过判断单元格格式的id解决，id值为58）
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
                double value = cell.getNumericCellValue();
                Date date = DateUtil.getJavaDate(value);
                result = simpleDateFormat.format(date);
            } else {
                double value = cell.getNumericCellValue();
                CellStyle cellStyle = cell.getCellStyle();
                DecimalFormat decimalFormat = new DecimalFormat();
                String temp = cellStyle.getDataFormatString();
                // 单元格设置成常规
                if (temp.equals("General")) {
                    decimalFormat.applyPattern("#");
                }
                result = decimalFormat.format(value);
            }
        }else if (cell.getCellTypeEnum() == CellType.STRING){
            result = cell.getStringCellValue().toString();
        }else if (cell.getCellTypeEnum() == CellType.BLANK) {
            result = "";
        }else {
            result = "";
        }
//        switch(cell.getCellType()) {
//            case Cell.CELL_TYPE_NUMERIC:    // 数字类型
//                if (HSSFDateUtil.isCellDateFormatted(cell)) {
//                    SimpleDateFormat simpleDateFormat = null;
//                    // 时间
//                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
//                        simpleDateFormat = new SimpleDateFormat("HH:mm");
//                    } else {
//                        // 日期
//                        simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
//                    }
//                    Date date = cell.getDateCellValue();
//                    result = simpleDateFormat.format(date);
//                } else if (cell.getCellStyle().getDataFormat() == 58) {
//                    // 处理自定义日期格式：m月d日（通过判断单元格格式的id解决，id值为58）
//                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
//                    double value = cell.getNumericCellValue();
//                    Date date = DateUtil.getJavaDate(value);
//                    result = simpleDateFormat.format(date);
//                } else {
//                    double value = cell.getNumericCellValue();
//                    CellStyle cellStyle = cell.getCellStyle();
//                    DecimalFormat decimalFormat = new DecimalFormat();
//                    String temp = cellStyle.getDataFormatString();
//                    // 单元格设置成常规
//                    if (temp.equals("General")) {
//                        decimalFormat.applyPattern("#");
//                    }
//                    result = decimalFormat.format(value);
//                }
//                break;
//            case Cell.CELL_TYPE_STRING: // 字符串
//                result = cell.getStringCellValue().toString();
//                break;
//            case Cell.CELL_TYPE_BLANK:
//                result = "";
//                break;
//            default:
//                result = "";
//                break;
//        }
        return result;
    }

    /**
     * 合并单元格
     * @param sheet
     * @return
     */
    private static Map<String, String>[] getRowSpanColSpan(Sheet sheet) {
        Map<String, String> map0 = new HashMap<String, String>();
        Map<String, String> map1 = new HashMap<String, String>();
        // 获取合并后的单元格数量
        int mergeNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        for (int i = 0; i < mergeNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            map0.put(topRow + "," + topCol, bottomRow + "," +bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, "");
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = {map0, map1};
        return map;
    }

    static String[] bordesr = { "border-top:", "border-right:", "border-bottom:", "border-left:" };
    static String[] borderStyles = { "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid", "solid", "solid", "solid", "solid" };

    /**
     * 处理单元格样式
     * @param wb
     * @param sheet
     * @param cell
     * @param sb
     */
    private static void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuffer sb) {
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
//            short alignment = cellStyle.getAlignment();
            short alignment = cellStyle.getAlignmentEnum().getCode();
            // 单元格内容的水平对齐方式
            sb.append("align='" + convertAlignToHtml(alignment) + "' ");
//            short verticalAlignment = cellStyle.getVerticalAlignment();
            short verticalAlignment = cellStyle.getVerticalAlignmentEnum().getCode();
            // 单元格中内容的垂直排列方式
            sb.append("valign='" + convertVerticalAlignToHtml(verticalAlignment) + "' ");

            if (wb instanceof XSSFWorkbook) {

                XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
//                short boldWeight = xf.getBoldweight();
                short boldWeight = (short) (xf.getBold() ? 700 : 400);
                sb.append("style='");
                sb.append("font-weight:" + boldWeight + ";");   // 字体加粗
                sb.append("font-size: " + xf.getFontHeight() / 2 + "%;");   // 字体大小
                int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                sb.append("width:" + columnWidth + "px;");

                XSSFColor xc = xf.getXSSFColor();
                if (xc != null && !"".equals(xc)) {
                    sb.append("color:#" + xc.getARGBHex().substring(2) + ";");  // 字体颜色
                }

                XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                if (bgColor != null && !"".equals(bgColor)) {
                    sb.append("background-color:#" + bgColor.getARGBHex().substring(2) + ";");  // 背景颜色
                }
                sb.append(getBorderStyle(0, cellStyle.getBorderTop(), ((XSSFCellStyle) cellStyle).getTopBorderXSSFColor()));
                sb.append(getBorderStyle(1, cellStyle.getBorderRight(), ((XSSFCellStyle) cellStyle).getRightBorderXSSFColor()));
                sb.append(getBorderStyle(2, cellStyle.getBorderBottom(), ((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor()));
                sb.append(getBorderStyle(3, cellStyle.getBorderLeft(), ((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor()));

            } else if (wb instanceof HSSFWorkbook) {

                HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
//                short boldWeight = hf.getBoldweight();
                short boldWeight = (short) (hf.getBold() ? 700 : 400);
                short fontColor = hf.getColor();
                sb.append("style='");
                HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
                HSSFColor hc = palette.getColor(fontColor);
                sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                sb.append("font-size: " + hf.getFontHeight() / 2 + "%;"); // 字体大小
                String fontColorStr = convertToStardColor(hc);
                if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
                    sb.append("color:" + fontColorStr + ";"); // 字体颜色
                }
                int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                sb.append("width:" + columnWidth + "px;");
                short bgColor = cellStyle.getFillForegroundColor();
                hc = palette.getColor(bgColor);
                String bgColorStr = convertToStardColor(hc);
                if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                    sb.append("background-color:" + bgColorStr + ";"); // 背景颜色
                }
                sb.append(getBorderStyle(palette, 0, cellStyle.getBorderTop(), cellStyle.getTopBorderColor()));
                sb.append(getBorderStyle(palette, 1, cellStyle.getBorderRight(), cellStyle.getRightBorderColor()));
                sb.append(getBorderStyle(palette, 3, cellStyle.getBorderLeft(), cellStyle.getLeftBorderColor()));
                sb.append(getBorderStyle(palette, 2, cellStyle.getBorderBottom(), cellStyle.getBottomBorderColor()));
            }

            sb.append("' ");
        }
    }

    /**
     * 垂直对齐方式
     * @param verticalAlignment
     * @return
     */
    private static String convertVerticalAlignToHtml(short verticalAlignment) {
        String valign = "middle";
//        switch (verticalAlignment) {
//            case CellStyle.VERTICAL_BOTTOM:
//                valign = "bottom";
//                break;
//            case CellStyle.VERTICAL_CENTER:
//                valign = "center";
//                break;
//            case CellStyle.VERTICAL_TOP:
//                valign = "top";
//                break;
//            default:
//                break;
//        }
        if (VerticalAlignment.BOTTOM.getCode() == verticalAlignment) {
            valign = "bottom";
        }else if (VerticalAlignment.CENTER.getCode() == verticalAlignment) {
            valign = "center";
        }else if (VerticalAlignment.TOP.getCode() == verticalAlignment) {
            valign = "top";
        }
        return valign;
    }

    /**
     * 水平对齐方式
     * @param alignment
     * @return
     */
    private static String convertAlignToHtml(short alignment) {
        String align = "left";
//        switch (alignment) {
//            case CellStyle.ALIGN_LEFT:
//                align = "left";
//                break;
//            case CellStyle.ALIGN_CENTER:
//                align = "center";
//                break;
//            case CellStyle.ALIGN_RIGHT:
//                align = "right";
//                break;
//            default:
//                break;
//        }
        if (HorizontalAlignment.LEFT.getCode() == alignment) {
            align = "left";
        }else if (HorizontalAlignment.CENTER.getCode() == alignment) {
            align = "center";
        }else if (HorizontalAlignment.RIGHT.getCode() == alignment) {
            align = "right";
        }
        return align;
    }

    private static String getBorderStyle(int b, short s, XSSFColor xc) {
        if (s == 0)
            return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
        ;
        if (xc != null && !"".equals(xc)) {
            String borderColorStr = xc.getARGBHex();// t.getARGBHex();
            borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr.substring(2);
            return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";
        }

        return "";
    }

    private static String getBorderStyle(HSSFPalette palette, int b, short s, short t) {
        if (s == 0)
            return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
        ;
        String borderColorStr = convertToStardColor(palette.getColor(t));
        borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr;
        return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";

    }

    private static String convertToStardColor(HSSFColor hc) {
        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
//            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
            if (HSSFColor.HSSFColorPredefined.AUTOMATIC.getIndex() == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }
        return sb.toString();
    }

    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }
}
