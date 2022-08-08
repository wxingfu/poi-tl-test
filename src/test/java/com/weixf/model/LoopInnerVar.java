package com.weixf.model;

public class LoopInnerVar {

    private String varName;
    private String varType;
    private String varDesc;

    public LoopInnerVar() {
    }

    public LoopInnerVar(String varName, String varType, String varDesc) {
        this.varName = varName;
        this.varType = varType;
        this.varDesc = varDesc;
    }

    public String getVarName() {
        return varName;
    }

    public void setVarName(String varName) {
        this.varName = varName;
    }

    public String getVarType() {
        return varType;
    }

    public void setVarType(String varType) {
        this.varType = varType;
    }

    public String getVarDesc() {
        return varDesc;
    }

    public void setVarDesc(String varDesc) {
        this.varDesc = varDesc;
    }

    @Override
    public String toString() {
        return "LoopInnerVar{" +
                "varName='" + varName + '\'' +
                ", varType='" + varType + '\'' +
                ", varDesc='" + varDesc + '\'' +
                '}';
    }
}
