package com.excel.entity;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-11-02-22:10
 */

public class ErrRowUserData {

    private String username;

    private String name;

    private String phone;

    private String email;

    private String expireTime;

    private String pwdPolicy;

    private String limitPwd;

    private String limitMac;

    private String limitUpSpeed;

    private String limitDownSpeed;

    private String desc;

    public ErrRowUserData() {
    }

    public ErrRowUserData(String username, String name, String phone, String email, String expireTime, String pwdPolicy, String limitPwd, String limitMac, String limitUpSpeed, String limitDownSpeed, String desc) {
        this.username = username;
        this.name = name;
        this.phone = phone;
        this.email = email;
        this.expireTime = expireTime;
        this.pwdPolicy = pwdPolicy;
        this.limitPwd = limitPwd;
        this.limitMac = limitMac;
        this.limitUpSpeed = limitUpSpeed;
        this.limitDownSpeed = limitDownSpeed;
        this.desc = desc;
    }

    @Override
    public String toString() {
        return "ErrRowUserData{" +
                "username='" + username + '\'' +
                ", name='" + name + '\'' +
                ", phone='" + phone + '\'' +
                ", email='" + email + '\'' +
                ", expireTime='" + expireTime + '\'' +
                ", pwdPolicy='" + pwdPolicy + '\'' +
                ", limitPwd='" + limitPwd + '\'' +
                ", limitMac='" + limitMac + '\'' +
                ", limitUpSpeed='" + limitUpSpeed + '\'' +
                ", limitDownSpeed='" + limitDownSpeed + '\'' +
                ", desc='" + desc + '\'' +
                '}';
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

    public String getLimitPwd() {
        return limitPwd;
    }

    public void setLimitPwd(String limitPwd) {
        this.limitPwd = limitPwd;
    }

    public String getLimitMac() {
        return limitMac;
    }

    public void setLimitMac(String limitMac) {
        this.limitMac = limitMac;
    }

    public String getLimitUpSpeed() {
        return limitUpSpeed;
    }

    public void setLimitUpSpeed(String limitUpSpeed) {
        this.limitUpSpeed = limitUpSpeed;
    }

    public String getLimitDownSpeed() {
        return limitDownSpeed;
    }

    public void setLimitDownSpeed(String limitDownSpeed) {
        this.limitDownSpeed = limitDownSpeed;
    }

    public String getPwdPolicy() {
        return pwdPolicy;
    }

    public void setPwdPolicy(String pwdPolicy) {
        this.pwdPolicy = pwdPolicy;
    }

    public String getExpireTime() {
        return expireTime;
    }

    public void setExpireTime(String expireTime) {
        this.expireTime = expireTime;
    }
}
