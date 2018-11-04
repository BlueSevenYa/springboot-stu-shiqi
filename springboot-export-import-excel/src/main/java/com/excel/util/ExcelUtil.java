package com.excel.util;

import com.excel.vo.ResultResp;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-11-04-15:00
 */

public class ExcelUtil {

    private static final Logger log= LoggerFactory.getLogger(ExcelUtil.class);

    public static final Integer notHeadExcelNum = 1; //非标题行行数

    public static final String[] headName=new String[]{"*用户名","姓名","电话","邮箱","简介",
            "限制密码数","限制Mac数","限制上行速度","限制下行速度","*密码策略",};


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

    }

    /**
     * 设置单元格样式
     * @param workbook
     * @return
     */
    public static CellStyle setCellStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(HSSFColor.YELLOW.index);
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
