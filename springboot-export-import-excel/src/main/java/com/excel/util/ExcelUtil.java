package com.excel.util;

import com.excel.entity.ErrExcelUserData;
import com.excel.entity.ErrRowUserData;
import com.excel.entity.User;
import com.excel.vo.ResultResp;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-11-04-15:00
 */

public class ExcelUtil {

    private static final Logger log= LoggerFactory.getLogger(ExcelUtil.class);

    public static final Integer notHeadExcelNum = 1; //非标题行行数

    public static final String[] headName=new String[]{"*用户名","姓名","电话","邮箱","过期时间","*密码策略",
            "限制密码数","限制Mac数","限制上行速度","限制下行速度","备注"};

    public static final String regUserName = "^[0-9a-z]{4,8}$"; // 用户名正则


    /**
     * 通过url来获取网络文件的io
     * @param fileUrl
     * @return
     */
    public static ResultResp<Object> getInputStreamByUrl(String fileUrl){
        ResultResp<Object> respIn=new ResultResp<>();
        InputStream inputStream=null;
        URL url = null;
        try {
            url = new URL(fileUrl);
        } catch (MalformedURLException e) {
            respIn.setCode(0);
            respIn.setMessage("MalformedURLException");
            return respIn;
        }
        HttpURLConnection conn = null;
        try {
            conn = (HttpURLConnection)url.openConnection();
        } catch (IOException e) {
            respIn.setCode(0);
            respIn.setMessage("IOException");
            return respIn;
        }
        //设置超时间为3秒
        conn.setConnectTimeout(3*1000);
        //防止屏蔽程序抓取而返回403错误
        conn.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)");

        //得到输入流
        try {
            inputStream = conn.getInputStream();
        } catch (IOException e) {
            respIn.setCode(0);
            respIn.setMessage("IOException");
            return respIn;
        }
        respIn.setCode(1);
        respIn.setData(inputStream);
        return respIn;
    }

    /**
     * 通过io获取workbook
     * @param inputStream
     * @return
     */
    public static Workbook getWorkBookByIo(InputStream inputStream){
        Workbook workbook=null;
        try {
            workbook= WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    /**
     * 校验excel模板合法性
     * @param workbook
     * @param notHeadExcelNum
     * @param headName
     * @return
     */
    public static boolean checkTemplateRight(Workbook workbook,Integer notHeadExcelNum,String[] headName){
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(notHeadExcelNum);
        if(row == null){
            return false;
        }
        int firstCellNum = row.getFirstCellNum();
        int lastCellNum = row.getLastCellNum();
        log.info(firstCellNum + " "+ lastCellNum);
        if(firstCellNum != 0){
            return false;
        }
        if(lastCellNum != headName.length){
            return false;
        }
        List<String> headNameStr=new ArrayList<>();
        for(int i=firstCellNum;i<lastCellNum;i++){
            Cell cell = row.getCell(i);
            headNameStr.add(cell.getStringCellValue());
        }
        if(checkListEqualsArr(headNameStr,headName)){
            return true;
        }else{
            return false;
        }
    }

    /**
     * 判断list string型和数组是否相等
     * @param list
     * @param arr
     * @return
     */
    public static boolean checkListEqualsArr(List<String> list,String[] arr){
        if(list.size() != arr.length){
            return false;
        }
        String[] temp=new String[list.size()];
        list.toArray(temp);
        if(Arrays.equals(temp,arr)){
            return true;
        }else{
            return false;
        }
    }

    public static void importExcelDeal(Workbook workbook){
        List<User> succUsers=new ArrayList<>();
        List<ErrExcelUserData> errUsers=new ArrayList<>();
        Set<String> userNameSet =new HashSet<>();

        Sheet sheet = workbook.getSheetAt(0);
        log.info(sheet.getLastRowNum()+"");
        for(int i=notHeadExcelNum+1;i<=sheet.getLastRowNum();i++){
            Row row= sheet.getRow(i);
            int firstCellNum= row.getFirstCellNum();
            int lastCellNum=row.getLastCellNum();
            int pCellNum=row.getPhysicalNumberOfCells();
            log.info(firstCellNum+ " " +lastCellNum +" " + pCellNum);
            /*for(int j=0;j<headName.length;j++){
                Cell cell=row.getCell(j);
                if(cell == null){
                    log.info("null");
                }else {
                    log.info(cell.getStringCellValue());
                }
            }*/

            getUserForRow(row,succUsers,errUsers,userNameSet);
        }

        if(errUsers.size() != 0){
            errUsers.forEach(errUser->{
                log.info(errUser.toString());
            });
        }
        log.info("end =================end errUsers");
        if(succUsers.size() != 0){
            succUsers.forEach(succUser->{
                log.info(succUser.toString());
            });
        }
        log.info("end =================end succUsers");
        userNameSet.forEach(unames->{
            log.info(unames);
        });
    }


    public static boolean getUserForRow(Row row,List<User> succUsers,List<ErrExcelUserData> errUsers,Set<String> userNameSet){
        Cell userNameCell = row.getCell(0);
        String originUserName = getCellValueByType(userNameCell);
        log.info(originUserName);

        Cell nameCell = row.getCell(1);
        String originName = getCellValueByType(nameCell);
        log.info(originName);

        Cell phoneCell = row.getCell(2);
        String originPhone = getCellValueByType(phoneCell);
        log.info(originPhone);

        Cell emailCell = row.getCell(3);
        String originEmail = getCellValueByType(emailCell);
        log.info(originEmail);

        Cell expireTimeCell = row.getCell(4);
        String originExpireTime = getCellValueByType(expireTimeCell);
        log.info(originExpireTime);

        Cell pwdPolicyCell = row.getCell(5);
        String originPwdPolicy = getCellValueByType(pwdPolicyCell);
        log.info(originPwdPolicy);

        Cell limitPwdCell = row.getCell(6);
        String originLimitPwd = getCellValueByType(limitPwdCell);
        log.info(originLimitPwd);

        Cell limitMacCell = row.getCell(7);
        String originLimitMac = getCellValueByType(limitMacCell);
        log.info(originLimitMac);

        Cell limitUpSpeedCell = row.getCell(8);
        String originLimitUpSpeed = getCellValueByType(limitUpSpeedCell);
        log.info(originLimitUpSpeed);

        Cell limitDownSpeedCell = row.getCell(9);
        String originLimitDownSpeed = getCellValueByType(limitDownSpeedCell);
        log.info(originLimitDownSpeed);

        Cell descCell = row.getCell(10);
        String originDesc = getCellValueByType(descCell);
        log.info(originDesc);

        boolean isHaveErrData = false;

        if(userNameCell == null){
            ErrRowUserData errRowUserData=new ErrRowUserData(originUserName,originName,originPhone,originEmail,originExpireTime,
                    originPwdPolicy,originLimitPwd,originLimitMac,originLimitUpSpeed,originLimitDownSpeed,originDesc);
            addErrRowUserToList(errUsers,errRowUserData,0,"字段不能为空","username");
            log.info("username maybe is null");
            isHaveErrData=true;
        }else if(!checkCellTypeIsStringOrNum(userNameCell)){
            ErrRowUserData errRowUserData=new ErrRowUserData(originUserName,originName,originPhone,originEmail,originExpireTime,
                    originPwdPolicy,originLimitPwd,originLimitMac,originLimitUpSpeed,originLimitDownSpeed,originDesc);
            addErrRowUserToList(errUsers,errRowUserData,0,"字段类型错误","username");
            log.info("username maybe is wrong type");
            isHaveErrData=true;
        }else if(!originUserName.matches(regUserName)){
            ErrRowUserData errRowUserData=new ErrRowUserData(originUserName,originName,originPhone,originEmail,originExpireTime,
                    originPwdPolicy,originLimitPwd,originLimitMac,originLimitUpSpeed,originLimitDownSpeed,originDesc);
            addErrRowUserToList(errUsers,errRowUserData,0,"字段格式错误","username");
            log.info("username maybe is wrong reg");
            isHaveErrData=true;
        }else if(userNameSet.contains(originUserName)){
            //进行excel用户名重复校验
            ErrRowUserData errRowUserData=new ErrRowUserData(originUserName,originName,originPhone,originEmail,originExpireTime,
                    originPwdPolicy,originLimitPwd,originLimitMac,originLimitUpSpeed,originLimitDownSpeed,originDesc);
            addErrRowUserToList(errUsers,errRowUserData,0,"表格中用户名重复","username");
            log.info("username maybe is repeat in excel");
            isHaveErrData=true;
        }else if(true){
            // 进行数据库层面用户名重复校验

            log.info("username maybe is repeat in db");
        }
        // 一旦校验到错误，就不在接着往下校验
        if(isHaveErrData){
            return isHaveErrData;
        }

        User user = new User();
        user.setName("success" +row.getRowNum());
        succUsers.add(user);
        userNameSet.add(originUserName);

        return isHaveErrData;
    }


    public static boolean checkCellTypeIsStringOrNum(Cell cell){
        if(cell.getCellTypeEnum()!= CellType.STRING && cell.getCellTypeEnum() != CellType.NUMERIC){
            return false;
        }else{
            return true;
        }
    }

    public static void addErrRowUserToList(List<ErrExcelUserData> errUsers,ErrRowUserData errRowUserData,Integer cellNum,String errMsg,String property){
        ErrExcelUserData errExcelUserData=new ErrExcelUserData(errRowUserData,cellNum,errMsg,property);
        errUsers.add(errExcelUserData);
    }


    public static String getCellValueByType(Cell cell){
        String value="";
        if(cell == null){
            log.info("string value is null");
        }else {
            switch (cell.getCellTypeEnum()) {
                case NUMERIC: // 数字
                    //如果为时间格式的内容
                    if (DateUtil.isCellDateFormatted(cell)) {
                        //注：format格式 yyyy-MM-dd hh:mm:ss 中小时为12小时制，若要24小时制，则把小h变为H即可，yyyy-MM-dd HH:mm:ss
                        Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                        value = sdf.format(date).toString();
                        break;
                    } else {
                        value = String.valueOf(cell.getNumericCellValue());
                        DecimalFormat df = new DecimalFormat("#.#########");
                        value=df.format(Double.valueOf(value));
                    }
                    break;

                case STRING: // 字符串
                    value = cell.getStringCellValue();
                    break;
                case BOOLEAN: // Boolean
                    value = cell.getBooleanCellValue() + "";
                    break;

                case FORMULA: // 公式
                    value = cell.getCellFormula() + "";
                    break;
                case BLANK: // 空值
                    value = "";
                    break;
                case ERROR: // 故障
                    value = "非法字符";
                    break;
                default:
                    value = "未知类型";
                    break;
            }
        }
        return value;
    }

    /**
     * 设置单元格样式
     * @param workbook
     * @return
     */
    public static CellStyle setCellStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.YELLOW.index);
        return style;
    }

    /**
     * 设置单元格批注
     * @param p
     * @return
     */
    public static HSSFComment setCellComment(HSSFPatriarch p,String message){
        //获取批注对象
        //(int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2)
        //前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
        HSSFComment comment=p.createComment(new HSSFClientAnchor(0,0,0,0,(short)3,3,(short)5,6));
        //输入批注信息
        comment.setString(new HSSFRichTextString(message));
        //添加作者
        comment.setAuthor("admin");
        comment.setFillColor(244,244,88);
        return comment;
    }
}
