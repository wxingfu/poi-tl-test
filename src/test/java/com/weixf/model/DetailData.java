package com.weixf.model;

import com.deepoove.poi.data.RowRenderData;

import java.util.List;

public class DetailData {

    private List<RowRenderData> goods;

    private List<RowRenderData> labors;

    public List<RowRenderData> getGoods() {
        return goods;
    }

    public void setGoods(List<RowRenderData> goods) {
        this.goods = goods;
    }

    public List<RowRenderData> getLabors() {
        return labors;
    }

    public void setLabors(List<RowRenderData> labors) {
        this.labors = labors;
    }
}
