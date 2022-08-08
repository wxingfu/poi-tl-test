package com.weixf.model;

import com.deepoove.poi.data.TextRenderData;

public class LoopPlugin {

    private String pluginName;
    private String pluginDesc;
    private TextRenderData remark;

    public LoopPlugin() {
    }

    public LoopPlugin(String pluginName, String pluginDesc, TextRenderData remark) {
        this.pluginName = pluginName;
        this.pluginDesc = pluginDesc;
        this.remark = remark;
    }

    public String getPluginName() {
        return pluginName;
    }

    public void setPluginName(String pluginName) {
        this.pluginName = pluginName;
    }

    public String getPluginDesc() {
        return pluginDesc;
    }

    public void setPluginDesc(String pluginDesc) {
        this.pluginDesc = pluginDesc;
    }

    public TextRenderData getRemark() {
        return remark;
    }

    public void setRemark(TextRenderData remark) {
        this.remark = remark;
    }

    @Override
    public String toString() {
        return "LoopPlugin{" +
                "pluginName='" + pluginName + '\'' +
                ", pluginDesc='" + pluginDesc + '\'' +
                ", remark='" + remark + '\'' +
                '}';
    }
}
