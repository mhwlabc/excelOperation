package com.example.excel;

import com.alibaba.fastjson.JSON;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.lang.Exception;

public class excel {

    public static List<Map<String, String>> readXlsxMap(String path) {
        try {
            InputStream is = new FileInputStream(path);
            XSSFWorkbook xssWorkbook = new XSSFWorkbook(is);
            List<Map<String, String>> result = new ArrayList<Map<String, String>>();
            for (XSSFSheet xssfSheet : xssWorkbook) {
                if (xssfSheet == null)
                    continue;
                XSSFRow keyRow = xssfSheet.getRow(0);
                for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                    XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                    int minCollx = xssfRow.getFirstCellNum();
                    int maxCollx = xssfRow.getLastCellNum();
                    Map<String, String> rowMap = new HashMap<String, String>();
                    for (int collx = minCollx; collx < maxCollx; collx++) {
                        XSSFCell xssfCell = xssfRow.getCell(collx);
                        rowMap.put(keyRow.getCell(collx).toString().trim(), getStringVal(xssfCell).trim());
                    }
                    result.add(rowMap);
                }
            }
            return result;
        } catch (Exception e) {
            List<Map<String, String>> map = new LinkedList<>();
            return map;
        }
    }

    private static String getStringVal(Cell cell) {
        if (cell == null) {
            return "";
        } else {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                case Cell.CELL_TYPE_FORMULA:
                    return cell.getCellFormula().replace("\"", "");
                case Cell.CELL_TYPE_NUMERIC:
                    if(HSSFDateUtil.isCellDateFormatted(cell)){
                        SimpleDateFormat sdf = null;
                        short i=cell.getCellStyle().getDataFormat();
                        //日期自定义类型特殊
                        if (cell.getCellStyle().getDataFormat() ==176){
                            sdf = new SimpleDateFormat("yyyy/MM/dd hh:mm:ss");
                        }else if (cell.getCellStyle().getDataFormat() ==22){
                            sdf = new SimpleDateFormat("yyyy/MM/dd hh:mm");
                        } else {
                            sdf = new SimpleDateFormat("yyyy/MM/dd");
                        }
                        Date date = cell.getDateCellValue();
                        return sdf.format(date);
                    }else {
                        cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                        return cell.getStringCellValue();
                    }
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue();
                default:
                    return StringUtils.EMPTY;
            }
        }
    }
}
