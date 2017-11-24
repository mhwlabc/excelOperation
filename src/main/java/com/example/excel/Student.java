package com.example.excel;
import java.util.Date;

public class Student{
    private String code;
    private String number;
    private String amount;
    private String order;
    private String liushui;
    private Date payDate;
    private String platform;
    public Student() {}

    public Student(String code,String number,String amount,String order,String liushui,Date payDate,String platform) {
        this.code = code;
        this.number = number;
        this.amount = amount;
        this.order = order;
        this.liushui = liushui;
        this.platform = platform;
        this.payDate = payDate;
    }

    public String getCode() {
        return code;
    }

    public String getNumber() {
        return number;
    }

    public String getAmount() {
        return amount;
    }

    public String getOrder() {
        return order;
    }

    public String getLiushui() {
        return liushui;
    }

    public Date getPayDate() {
        return payDate;
    }

    public String getPlatform() {
        return platform;
    }

}
