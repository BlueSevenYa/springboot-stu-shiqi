package com.excel.util;

import com.excel.entity.ErrExcelUserData;
import com.excel.entity.ErrRowUserData;
import com.excel.entity.User;
import com.excel.vo.ResultResp;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.swing.filechooser.FileSystemView;
import java.io.*;
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

    public static final String[] tipsMsg = new String[]{
            "*代表必填项。",
            "1. msg1",
            "2. msg2",
            "3. msg3"
    };


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

    /**
     * 处理导入逻辑
     * @param workbook
     */
    public static void importExcelDeal(Workbook workbook){
        List<User> succUsers=new ArrayList<>();
        List<ErrExcelUserData> errUsers=new ArrayList<>();
        Set<String> userNameSet =new HashSet<>();

        Sheet sheet = workbook.getSheetAt(0);
        log.info(sheet.getLastRowNum()+"");
        for(int i=notHeadExcelNum+1;i<=sheet.getLastRowNum();i++){
            Row row= sheet.getRow(i);
            log.info(row.getFirstCellNum()+ " " +row.getLastCellNum() +" " + row.getPhysicalNumberOfCells());

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

        exportExcelDeal(errUsers);
    }


    public static void exportExcelDeal(List<ErrExcelUserData> errUsers){
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("错误数据");
        //创建绘图对象
        HSSFPatriarch p=sheet.createDrawingPatriarch();

        CellRangeAddress region = new CellRangeAddress(0,0,0,10);
        sheet.addMergedRegion(region);

        Row firstRow = sheet.createRow(0);
        firstRow.setHeightInPoints((short) (80));

        Font firstRowFont = workbook.createFont();
        firstRowFont.setFontName("宋体");
        firstRowFont.setFontHeightInPoints((short) 11);

        CellStyle firstRowStyle = workbook.createCellStyle();
        firstRowStyle.setWrapText(true); // 设置自动换行
        firstRowStyle.setAlignment(HorizontalAlignment.LEFT);
        firstRowStyle.setVerticalAlignment(VerticalAlignment.TOP);
        firstRowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        firstRowStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
        firstRowStyle.setFont(firstRowFont);

        Cell firstRowCell = firstRow.createCell(0);
        StringBuilder sb = new StringBuilder();
        for (String s : tipsMsg) {
            sb.append(s).append("\r\n");
        }
        HSSFRichTextString richString = new HSSFRichTextString(sb.toString().trim());
        Font redFont = workbook.createFont();
        redFont.setColor(IndexedColors.RED.index);
        redFont.setFontHeightInPoints((short) 11);
        redFont.setFontName("宋体");
        richString.applyFont(0,1,redFont);
        richString.applyFont(1,richString.length()-1,firstRowFont);
        firstRowCell.setCellStyle(firstRowStyle);
        firstRowCell.setCellValue(richString);


        Row secondRow =sheet.createRow(1);
        for(int i=0;i<headName.length;i++){
            Font font = workbook.createFont();
            font.setFontName("宋体");
            font.setFontHeightInPoints((short) 12);
            font.setBold(true);

            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.index);
            cellStyle.setFont(font);
            Cell cell=secondRow.createCell(i);
            cell.setCellValue(headName[i]);
            cell.setCellStyle(cellStyle);
        }



        //让列宽随着导出的列长自动适应
        for (int colNum = 0; colNum < headName.length; colNum++) {
            int columnWidth = sheet.getColumnWidth(colNum) / 256;
            for (int rowNum = 1; rowNum < sheet.getLastRowNum(); rowNum++) {
                HSSFRow currentRow;
                //当前行未被使用过
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }
                if (currentRow.getCell(colNum) != null) {
                    HSSFCell currentCell = currentRow.getCell(colNum);
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            if(colNum == 0){
                sheet.setColumnWidth(colNum, (columnWidth-2) * 256);
            }else{
                sheet.setColumnWidth(colNum, (columnWidth+4) * 256);
            }
        }


        //创建输出流
        FileOutputStream out=null;
        FileSystemView fsv = FileSystemView.getFileSystemView();
        File com=fsv.getHomeDirectory();    //读取桌面路径
        String deskTopPath = com.getPath();
        String filePath = deskTopPath +"\\"+"errUserData.xls";
        try {
            out=new FileOutputStream(filePath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 处理校验每一行的数据
     * @param row
     * @param succUsers
     * @param errUsers
     * @param userNameSet
     * @return
     */
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
        int pwdPolicy = getNumValue(pwdPolicyCell);
        log.info(pwdPolicy + "is after changing");

        Cell limitPwdCell = row.getCell(6);
        String originLimitPwd = getCellValueByType(limitPwdCell);
        log.info(originLimitPwd);
        int limitPwd = getNumValue(limitPwdCell);
        log.info(limitPwd + "is after changing");

        Cell limitMacCell = row.getCell(7);
        String originLimitMac = getCellValueByType(limitMacCell);
        log.info(originLimitMac);
        int limitMac = getNumValue(limitMacCell);
        log.info(limitMac + "is after changing");

        Cell limitUpSpeedCell = row.getCell(8);
        String originLimitUpSpeed = getCellValueByType(limitUpSpeedCell);
        log.info(originLimitUpSpeed);
        int limitUpSpeed = getNumValue(limitUpSpeedCell);
        log.info(limitUpSpeed + "is after changing");

        Cell limitDownSpeedCell = row.getCell(9);
        String originLimitDownSpeed = getCellValueByType(limitDownSpeedCell);
        log.info(originLimitDownSpeed);
        int limitDownSpeed = getNumValue(limitDownSpeedCell);
        log.info(limitDownSpeed + "is after changing");

        Cell descCell = row.getCell(10);
        String originDesc = getCellValueByType(descCell);
        log.info(originDesc);

        boolean isHaveErrData = false;

        // 用户名校验
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

        // 密码策略校验
        if(pwdPolicyCell == null){
            ErrRowUserData errRowUserData=new ErrRowUserData(originUserName,originName,originPhone,originEmail,originExpireTime,
                    originPwdPolicy,originLimitPwd,originLimitMac,originLimitUpSpeed,originLimitDownSpeed,originDesc);
            addErrRowUserToList(errUsers,errRowUserData,5,"字段不能为空","pwdPolicy");
            log.info("pwdPolicy maybe is null");
            isHaveErrData=true;
        }else if(!checkCellTypeIsNum(pwdPolicyCell)){
            ErrRowUserData errRowUserData=new ErrRowUserData(originUserName,originName,originPhone,originEmail,originExpireTime,
                    originPwdPolicy,originLimitPwd,originLimitMac,originLimitUpSpeed,originLimitDownSpeed,originDesc);
            addErrRowUserToList(errUsers,errRowUserData,5,"字段类型错误","pwdPolicy");
            log.info("pwdPolicy maybe wrong type");
            isHaveErrData=true;
        }else if(originPwdPolicy.indexOf(".") != -1){
            ErrRowUserData errRowUserData=new ErrRowUserData(originUserName,originName,originPhone,originEmail,originExpireTime,
                    originPwdPolicy,originLimitPwd,originLimitMac,originLimitUpSpeed,originLimitDownSpeed,originDesc);
            addErrRowUserToList(errUsers,errRowUserData,5,"字段格式错误","pwdPolicy");
            log.info("pwdPolicy maybe wrong geshi");
            isHaveErrData=true;
        } else if (pwdPolicy <1 || pwdPolicy >3){
            ErrRowUserData errRowUserData=new ErrRowUserData(originUserName,originName,originPhone,originEmail,originExpireTime,
                    originPwdPolicy,originLimitPwd,originLimitMac,originLimitUpSpeed,originLimitDownSpeed,originDesc);
            addErrRowUserToList(errUsers,errRowUserData,5,"字段范围错误","pwdPolicy");
            log.info("pwdPolicy maybe wrong range");
            isHaveErrData=true;
        }
        if(isHaveErrData){
            return isHaveErrData;
        }

        User user = new User();
        user.setName("success" +row.getRowNum());
        succUsers.add(user);
        userNameSet.add(originUserName);

        return isHaveErrData;
    }


    public static int getNumValue(Cell cell){
        if(cell == null){
            return 0;
        }
        if(!checkCellTypeIsNum(cell)){
            return 0;
        }
        return (int) cell.getNumericCellValue();
    }

    /**
     * 判断Cell的类型是否是数字或者字符串
     * @param cell
     * @return
     */
    public static boolean checkCellTypeIsStringOrNum(Cell cell){
        if(cell.getCellTypeEnum()!= CellType.STRING && cell.getCellTypeEnum() != CellType.NUMERIC){
            return false;
        }else{
            return true;
        }
    }


    /**
     * 判断Cell的类型是数字并且不是日期
     * @param cell
     * @return
     */
    public static boolean checkCellTypeIsNum(Cell cell){
        if(cell.getCellTypeEnum() == CellType.NUMERIC && !DateUtil.isCellDateFormatted(cell)){
            return true;
        }else{
            return false;
        }
    }

    /**
     * 封装添加错误RowUser到list
     * @param errUsers
     * @param errRowUserData
     * @param cellNum
     * @param errMsg
     * @param property
     */
    public static void addErrRowUserToList(List<ErrExcelUserData> errUsers,ErrRowUserData errRowUserData,Integer cellNum,String errMsg,String property){
        ErrExcelUserData errExcelUserData=new ErrExcelUserData(errRowUserData,cellNum,errMsg,property);
        errUsers.add(errExcelUserData);
    }


    /**
     * 得到excel 数据的原始值，转成String类型
     * @param cell
     * @return
     */
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
