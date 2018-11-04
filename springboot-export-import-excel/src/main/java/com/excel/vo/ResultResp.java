package com.excel.vo;

import java.io.Serializable;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-11-02-22:17
 */

public class ResultResp<T> implements Serializable{

    private Integer code;

    private String message;

    private T data;

    @Override
    public String toString() {
        return "ResultResp{" +
                "code=" + code +
                ", message='" + message + '\'' +
                ", data=" + data +
                '}';
    }

    public Integer getCode() {
        return code;
    }

    public void setCode(Integer code) {
        this.code = code;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }
}
