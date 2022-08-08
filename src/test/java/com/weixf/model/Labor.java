package com.weixf.model;

public class Labor {


    private String category;
    private int people;
    private int price;
    private int totalPrice;

    public Labor() {
    }

    public Labor(String category, int people, int price, int totalPrice) {
        this.category = category;
        this.people = people;
        this.price = price;
        this.totalPrice = totalPrice;
    }

    public String getCategory() {
        return category;
    }

    public void setCategory(String category) {
        this.category = category;
    }

    public int getPeople() {
        return people;
    }

    public void setPeople(int people) {
        this.people = people;
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

    @Override
    public String toString() {
        return "Labor{" +
                "category='" + category + '\'' +
                ", people=" + people +
                ", price=" + price +
                ", totalPrice=" + totalPrice +
                '}';
    }
}
