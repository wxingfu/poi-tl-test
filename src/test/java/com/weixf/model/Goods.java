package com.weixf.model;

import com.deepoove.poi.data.PictureRenderData;

public class Goods {

    private int count;
    private String name;
    private String desc;
    private int discount;
    private int tax;
    private int price;
    private int totalPrice;
    private PictureRenderData picture;

    public Goods() {
    }

    public Goods(int count, String name, String desc, int discount, int tax, int price, int totalPrice, PictureRenderData picture) {
        this.count = count;
        this.name = name;
        this.desc = desc;
        this.discount = discount;
        this.tax = tax;
        this.price = price;
        this.totalPrice = totalPrice;
        this.picture = picture;
    }

    public int getCount() {
        return count;
    }

    public void setCount(int count) {
        this.count = count;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }

    public int getDiscount() {
        return discount;
    }

    public void setDiscount(int discount) {
        this.discount = discount;
    }

    public int getTax() {
        return tax;
    }

    public void setTax(int tax) {
        this.tax = tax;
    }

    public int getPrice() {
        return price;
    }

    public void setPrice(int price) {
        this.price = price;
    }

    public int getTotalPrice() {
        return totalPrice;
    }

    public void setTotalPrice(int totalPrice) {
        this.totalPrice = totalPrice;
    }

    public PictureRenderData getPicture() {
        return picture;
    }

    public void setPicture(PictureRenderData picture) {
        this.picture = picture;
    }

    @Override
    public String toString() {
        return "Goods{" +
                "count=" + count +
                ", name='" + name + '\'' +
                ", desc='" + desc + '\'' +
                ", discount=" + discount +
                ", tax=" + tax +
                ", price=" + price +
                ", totalPrice=" + totalPrice +
                ", picture=" + picture +
                '}';
    }
}
