package com.weixf.model;

public class AddrModel {

    private String addr;

    public AddrModel() {
    }

    public AddrModel(String addr) {
        this.addr = addr;
    }

    public String getAddr() {
        return addr;
    }

    public void setAddr(String addr) {
        this.addr = addr;
    }

    @Override
    public String toString() {
        return "AddrModel{" +
                "addr='" + addr + '\'' +
                '}';
    }
}
