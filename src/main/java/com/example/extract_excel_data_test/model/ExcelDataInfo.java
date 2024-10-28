package com.example.extract_excel_data_test.model;


import lombok.Builder;

public class ExcelDataInfo {

    private String name;
    private String address;
    private String phone;
    private String email;
    private String comment;
    private String birthDay;
    private String job;

    @Builder
    public ExcelDataInfo(String name, String address, String phone, String email, String comment, String birthDay, String job) {
        this.name = name;
        this.address = address;
        this.phone = phone;
        this.email = email;
        this.comment = comment;
        this.birthDay = birthDay;
        this.job = job;
    }

    public ExcelDataInfo() {}



    public String getName() {
        return name;
    }

    public String getAddress() {
        return address;
    }

    public String getPhone() {
        return phone;
    }

    public String getEmail() {
        return email;
    }

    public String getComment() {
        return comment;
    }

    public String getBirthDay() {
        return birthDay;
    }

    public String getJob() {
        return job;
    }

    public void updateName(String name) {
        this.name = name;
    }

    public void updateAddress(String address) {
        this.address = address;
    }

    public void updatePhone(String phone) {
        this.phone = phone;
    }

    public void updateEmail(String email) {
        this.email = email;
    }

    public void updateComment(String comment) {
        this.comment = comment;
    }

    public void updateBirthDay(String birthDay) {
        this.birthDay = birthDay;
    }

    public void updateJob(String job) {
        this.job = job;
    }

    @Override
    public String toString() {
        return "ExcelDataInfo{" +
                "name='" + name + '\'' +
                ", address='" + address + '\'' +
                ", phone='" + phone + '\'' +
                ", email='" + email + '\'' +
                ", comment='" + comment + '\'' +
                ", birthDay='" + birthDay + '\'' +
                ", job='" + job + '\'' +
                '}';
    }
}
