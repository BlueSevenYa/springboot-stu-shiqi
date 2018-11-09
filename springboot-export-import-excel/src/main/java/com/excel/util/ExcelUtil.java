package com.excel.util;

import com.excel.entity.ErrExcelUserData;
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

    public static final String[] headName=new String[]{"*用户名","姓名","电话","邮箱","过期时间","*密码策略",
            "限制密码数","限制Mac数","限制上行速度","限制下行速度","简介"};




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

        Sheet sheet = workbook.getSheetAt(0);
        log.info(sheet.getLastRowNum()+"");
        for(int i=notHeadExcelNum+1;i<=sheet.getLastRowNum();i++){
            Row row= sheet.getRow(i);
            int firstCellNum= row.getFirstCellNum();
            int lastCellNum=row.getLastCellNum();
            int pCellNum=row.getPhysicalNumberOfCells();
            log.info(firstCellNum+ " " +lastCellNum +" " + pCellNum);
            log.info(row.getCell(0).getStringCellValue());
            for(int j=0;j<headName.length;j++){
                Cell cell=row.getCell(j);
                if(cell == null){
                    log.info("null");
                }else {
                    log.info(cell.getStringCellValue());
                }
            }


        }
    }


    public static void getUserForRow(Row row,List<User> succUsers,List<ErrExcelUserData> errUsers){
        Cell userNameCell = row.getCell(0);

        Cell nameCell = row.getCell(1);

        Cell phoneCell = row.getCell(2);

        Cell email = row.getCell(3);

        Cell expireTimeCell = row.getCell(4);

        Cell pwdPolicyCell = row.getCell(5);

        Cell limitPwdCell = row.getCell(6);

        Cell limitMacCell = row.getCell(7);

        Cell limitUpSpeedCell = row.getCell(8);

        Cell limitDownSpeedCell = row.getCell(9);

        Cell descCell = row.getCell(10);
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
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                        value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue())).toString();
                        break;
                    } else {
                        //value = new DecimalFormat("0").format(cell.getNumericCellValue());
                        value = String.valueOf(cell.getNumericCellValue());
                        DecimalFormat df = new DecimalFormat("#.#########");
                        value=df.format(Double.valueOf(value));
                    }
                    break;

                /*if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    //如果是date类型则 ，获取该cell的date值
                    Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                    SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    value = format.format(date);;
                }else {// 纯数字
                    BigDecimal big=new BigDecimal(cell.getNumericCellValue());
                    value = big.toString();
                    //解决1234.0  去掉后面的.0
                    if(null!=value&&!"".equals(value.trim())){
                        String[] item = value.split("[.]");
                        if(1<item.length&&"0".equals(item[1])){
                            value=item[0];
                        }
                    }
                }
                break;*/
                case STRING: // 字符串
                    value = cell.getStringCellValue();
                    break;
                case BOOLEAN: // Boolean
                    value = cell.getBooleanCellValue() + "";
                    break;
                /*value = " "+ cell.getBooleanCellValue();
                break;*/
                case FORMULA: // 公式
                    value = cell.getCellFormula() + "";
                    break;
                /*value = String.valueOf(cell.getNumericCellValue());
                if (value.equals("NaN")) {// 如果获取的数据值为非法值,则转换为获取字符串
                    value = cell.getStringCellValue().toString();
                }
                break;*/
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
