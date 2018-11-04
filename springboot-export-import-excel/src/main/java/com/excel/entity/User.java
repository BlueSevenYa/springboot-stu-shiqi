package com.excel.entity;

import java.util.Date;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-10-30-21:46
 */

public class User {

    private String id;

    private String uid;

    private String username;

    private String name;

    private String phone;

    private String email;

    private String desc;

    private Date createTime;

    private Date expireTime;

    private Integer limitPwd;

    private Integer limitMac;

    private Integer limitUpSpeed;

    private Integer limitDownSpeed;

    private Integer pwdPolicy;//密码策略类型

    @Override
    public String toString() {
        return "User{" +
                "id='" + id + '\'' +
                ", uid='" + uid + '\'' +
                ", username='" + username + '\'' +
                ", name='" + name + '\'' +
                ", phone='" + phone + '\'' +
                ", email='" + email + '\'' +
                ", desc='" + desc + '\'' +
                ", createTime=" + createTime +
                ", expireTime=" + expireTime +
                ", limitPwd=" + limitPwd +
                ", limitMac=" + limitMac +
                ", limitUpSpeed=" + limitUpSpeed +
                ", limitDownSpeed=" + limitDownSpeed +
                ", pwdPolicy=" + pwdPolicy +
                '}';
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getUid() {
        return uid;
    }

    public void setUid(String uid) {
        this.uid = uid;
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }

    public Date getCreateTime() {
        return createTime;
    }

    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }

    public Date getExpireTime() {
        return expireTime;
    }

    public void setExpireTime(Date expireTime) {
        this.expireTime = expireTime;
    }

    public Integer getLimitPwd() {
        return limitPwd;
    }

    public void setLimitPwd(Integer limitPwd) {
        this.limitPwd = limitPwd;
    }

    public Integer getLimitMac() {
        return limitMac;
    }

    public void setLimitMac(Integer limitMac) {
        this.limitMac = limitMac;
    }

    public Integer getLimitUpSpeed() {
        return limitUpSpeed;
    }

    public void setLimitUpSpeed(Integer limitUpSpeed) {
        this.limitUpSpeed = limitUpSpeed;
    }

    public Integer getLimitDownSpeed() {
        return limitDownSpeed;
    }

    public void setLimitDownSpeed(Integer limitDownSpeed) {
        this.limitDownSpeed = limitDownSpeed;
    }

    public Integer getPwdPolicy() {
        return pwdPolicy;
    }

    public void setPwdPolicy(Integer pwdPolicy) {
        this.pwdPolicy = pwdPolicy;
    }
}
