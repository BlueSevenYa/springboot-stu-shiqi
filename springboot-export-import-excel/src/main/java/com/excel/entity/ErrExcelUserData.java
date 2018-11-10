package com.excel.entity;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-11-02-22:13
 */

public class ErrExcelUserData {

    private ErrRowUserData errRowUserData;

    private Integer cellNum;

    private String errMsg;

    private String property;

    public ErrExcelUserData() {
    }

    public ErrExcelUserData(ErrRowUserData errRowUserData, Integer cellNum, String errMsg, String property) {
        this.errRowUserData = errRowUserData;
        this.cellNum = cellNum;
        this.errMsg = errMsg;
        this.property = property;
    }

    @Override
    public String toString() {
        return "ErrExcelUserData{" +
                "errRowUserData=" + errRowUserData +
                ", errMsg='" + errMsg + '\'' +
                ", property='" + property + '\'' +
                ", cellNum=" + cellNum +
                '}';
    }

    public ErrRowUserData getErrRowUserData() {
        return errRowUserData;
    }

    public void setErrRowUserData(ErrRowUserData errRowUserData) {
        this.errRowUserData = errRowUserData;
    }

    public String getErrMsg() {
        return errMsg;
    }

    public void setErrMsg(String errMsg) {
        this.errMsg = errMsg;
    }

    public String getProperty() {
        return property;
    }

    public void setProperty(String property) {
        this.property = property;
    }

    public Integer getCellNum() {
        return cellNum;
    }

    public void setCellNum(Integer cellNum) {
        this.cellNum = cellNum;
    }
}
