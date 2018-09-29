package com.drools.model;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-09-29-22:09
 */

public class Product {

    /*
     * 钻石
     */
    public static final String DIAMOND="0";

    /*
     * 黄金
     */
    public static final String GOLD="1";

    private String type;

    private Integer discount;

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public Integer getDiscount() {
        return discount;
    }

    public void setDiscount(Integer discount) {
        this.discount = discount;
    }
}
