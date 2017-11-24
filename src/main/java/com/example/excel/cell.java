package com.example.excel;

import com.alibaba.fastjson.JSON;
import org.apache.commons.lang3.time.DateUtils;

import java.util.Date;
import java.util.List;
import java.util.Map;

public class cell {
    public static void importRepayDetail(long planDetailId, int repayOrder, String url) throws Exception {

        List<Map<String, String>> readExcelMap = excel.readXlsxMap(url);
        for (Map<String, String> list : readExcelMap) {
            nullOrString(list.get("资产订单编号"));
            nullOrString(list.get("期次"));
            nullOrString(list.get("还款支付金额"));
            nullOrString(list.get("还款支付通道商户订单号"));
            nullOrString(list.get("还款支付通道商户流水号"));
            nullOrString(list.get("支付通道还款时间"));
            nullOrString(list.get("还款支付通道名"));

            System.out.println(list.get("资产订单编号").getClass());
            System.out.println(list.get("期次"));
            System.out.println(list.get("还款支付金额"));
            System.out.println(list.get("还款支付通道商户订单号"));
            System.out.println(list.get("还款支付通道商户流水号"));
            System.out.println(list.get("支付通道还款时间"));
            System.out.println(list.get("还款支付通道名"));
        }

        for (Map<String, String> list : readExcelMap) {
            Date date = new Date();
            try {
                date = DateUtils.parseDate(
                        list.get("支付通道还款时间"), "yyyy/MM/dd HH:mm"
                );
            } catch (Exception e) {
                throw new Exception("错了");
            }

        }


        System.out.println("HHHHHH");
    }

    //判断表中字段是否完整
    public static void nullOrString(String str) throws Exception {
        if (str == null) {
            throw new Exception("不能为空");
        }
    }

    public static void main(String args[]) throws Exception {
        importRepayDetail(123,3,"C:\\Users\\abczh\\Desktop\\test.xlsx");

        TestOne testOne = new TestOne();
        System.out.println("1:"+testOne.getTestOne());
        System.out.println("2:"+testOne.getTestTwo());
        System.out.println("3:"+testOne.getTestThree());
        System.out.println("4:"+testOne.getTestFour());
    }

}
