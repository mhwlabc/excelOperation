package com.example.excel;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class toExcel {
        /**
         * @功能：手工构建一个简单格式的Excel
         */

        private static List getStudent() throws Exception {

            List<Map<String, String>> readExcelMap = excel.readXlsxMap("C:\\Users\\abczh\\Desktop\\test.xlsx");

            List list = new ArrayList();

            for (Map<String, String> map : readExcelMap) {
                String code = (String) map.get("资产订单编号");
                String number = (String) map.get("期次");
                String amount = (String) map.get("还款支付金额");
                String order = (String) map.get("还款支付通道商户订单号");
                String liushui = (String) map.get("还款支付通道商户流水号");
                String payDate = (String) map.get("支付通道还款时间");
                String platform = (String) map.get("还款支付通道名");

                Date date = new Date();
                date = DateUtils.parseDate(
                        payDate, "yyyy/MM/dd HH:mm"
                );
                Student xls = new Student(code, number,amount,order,liushui,date,platform);
                list.add(xls);
            }
            System.out.println(readExcelMap.size());
            return list;
        }

        public static void main(String[] args) throws Exception
        {
            // 第一步，创建一个webbook，对应一个Excel文件
            HSSFWorkbook wb = new HSSFWorkbook();
            // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
            HSSFSheet sheet = wb.createSheet("还款明细");
            // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
            HSSFRow row = sheet.createRow((int) 0);
            // 第四步，创建单元格，并设置值表头 设置表头居中
            HSSFCellStyle style = wb.createCellStyle();
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

            HSSFCell cell = row.createCell((short) 0);
            cell.setCellValue("资产订单编号");
            cell.setCellStyle(style);
            cell = row.createCell((short) 1);
            cell.setCellValue("期次");
            cell.setCellStyle(style);
            cell = row.createCell((short) 2);
            cell.setCellValue("还款支付金额");
            cell.setCellStyle(style);
            cell = row.createCell((short) 3);
            cell.setCellValue("还款支付通道商户订单号");
            cell.setCellStyle(style);
            cell = row.createCell((short) 4);
            cell.setCellValue("还款支付通道商户流水号");
            cell.setCellStyle(style);
            cell = row.createCell((short) 5);
            cell.setCellValue("支付通道还款时间");
            cell.setCellStyle(style);
            cell = row.createCell((short) 6);
            cell.setCellValue("还款支付通道名");
            cell.setCellStyle(style);

            // 第五步，写入实体数据 实际应用中这些数据从数据库得到，
            List list = toExcel.getStudent();

            for (int i = 0; i < list.size(); i++) {
                row = sheet.createRow((int) i + 1);
                System.out.println("  "+list.size()+"---123");
                Student stu = (Student) list.get(i);
                // 第四步，创建单元格，并设置值
                row.createCell((short) 0).setCellValue(stu.getCode());
                row.createCell((short) 1).setCellValue(stu.getNumber());
                row.createCell((short) 2).setCellValue(stu.getAmount());
                row.createCell((short) 3).setCellValue(stu.getOrder());
                row.createCell((short) 4).setCellValue(stu.getLiushui());
                row.createCell((short) 4).setCellValue(stu.getPayDate());
                row.createCell((short) 6).setCellValue(stu.getPlatform());
            }

            // 第六步，将文件存到指定位置
            try
            {
                FileOutputStream fout = new FileOutputStream("C:\\Users\\abczh\\Desktop\\students.xls");
                wb.write(fout);
                fout.close();
            }
            catch (Exception e)
            {
                e.printStackTrace();
            }
        }
}
