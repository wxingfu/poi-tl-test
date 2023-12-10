package com.weixf.example;

import com.deepoove.poi.data.TableRenderData;
import com.deepoove.poi.expression.Name;

public class PaymentData {
    private TableRenderData order;
    private String NO;
    private String ID;
    private String taitou;
    private String consignee;
    @Name("detail_table")
    private DetailData detailTable;
    private String subtotal;
    private String tax;
    private String transform;
    private String other;
    private String unpay;
    private String total;

    public TableRenderData getOrder() {
        return this.order;
    }

    public void setOrder(TableRenderData order) {
        this.order = order;
    }

    public String getNO() {
        return this.NO;
    }

    public void setNO(String NO) {
        this.NO = NO;
    }

    public String getID() {
        return this.ID;
    }

    public void setID(String ID) {
        this.ID = ID;
    }

    public String getTaitou() {
        return this.taitou;
    }

    public void setTaitou(String taitou) {
        this.taitou = taitou;
    }

    public String getConsignee() {
        return this.consignee;
    }

    public void setConsignee(String consignee) {
        this.consignee = consignee;
    }

    public DetailData getDetailTable() {
        return this.detailTable;
    }

    public void setDetailTable(DetailData detailTable) {
        this.detailTable = detailTable;
    }

    public String getSubtotal() {
        return this.subtotal;
    }

    public void setSubtotal(String subtotal) {
        this.subtotal = subtotal;
    }

    public String getTax() {
        return this.tax;
    }

    public void setTax(String tax) {
        this.tax = tax;
    }

    public String getTransform() {
        return this.transform;
    }

    public void setTransform(String transform) {
        this.transform = transform;
    }

    public String getOther() {
        return this.other;
    }

    public void setOther(String other) {
        this.other = other;
    }

    public String getUnpay() {
        return this.unpay;
    }

    public void setUnpay(String unpay) {
        this.unpay = unpay;
    }

    public String getTotal() {
        return total;
    }

    public void setTotal(String total) {
        this.total = total;
    }

}
