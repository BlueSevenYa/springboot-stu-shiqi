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

    private String desc;

    private String limitPwd;

    private String limitMac;

    private String limitUpSpeed;

    private String limitDownSpeed;

    private String pwdPolicy;

    private String expireTime;

    @Override
    public String toString() {
        return "ErrRowUserData{" +
                "username='" + username + '\'' +
                ", name='" + name + '\'' +
                ", phone='" + phone + '\'' +
                ", email='" + email + '\'' +
                ", desc='" + desc + '\'' +
                ", limitPwd='" + limitPwd + '\'' +
                ", limitMac='" + limitMac + '\'' +
                ", limitUpSpeed='" + limitUpSpeed + '\'' +
                ", limitDownSpeed='" + limitDownSpeed + '\'' +
                ", pwdPolicy='" + pwdPolicy + '\'' +
                ", expireTime='" + expireTime + '\'' +
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
