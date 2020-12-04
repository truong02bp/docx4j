package com.demo.entities;

import java.util.Arrays;

public class Person {
    private String name;
    private String Address;
    private String company;
    private String phone;
    private byte[] image;

    public Person(String name, String address, String company, String phone) {
        this.name = name;
        Address = address;
        this.company = company;
        this.phone = phone;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAddress() {
        return Address;
    }

    public void setAddress(String address) {
        this.Address = address;
    }

    public String getCompany() {
        return company;
    }

    public void setCompany(String company) {
        this.company = company;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public byte[] getImage() {
        return image;
    }

    public void setImage(byte[] image) {
        this.image = image;
    }

    @Override
    public String toString() {
        return "name=" + name+';' +
                "Address=" + Address + ';' +
                "company=" + company + ';' +
                "phone=" + phone + ';';
    }
}
