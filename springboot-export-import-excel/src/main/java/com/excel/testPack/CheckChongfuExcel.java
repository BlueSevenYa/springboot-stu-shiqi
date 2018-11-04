package com.excel.testPack;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Created by
 *
 * @author=蓝十七
 * @on 2018-11-04-18:10
 */

public class CheckChongfuExcel {

    /**
     * 判定Excel中某列是否有重复数据
     * @param file 读取数据的源Excel
     * @param ignoreRows 读取数据忽略的行数，比如行头不需要读入 忽略的行数为1
     * @param column 需要判定的字段所在列的位置，比如需要判定的字段在第三列， column=2；注意，0是算第一列
     * @return 读出的Excel中数据的内容
     */

    public boolean judgeRepeatExcle(File file, int column, int ignoreRows){
        boolean flag=false;
        if(column>=0&&file.exists()){
	   /*实现excle的兼容读取*/
            Workbook wb = null;
            try {
                wb = new XSSFWorkbook(new FileInputStream(file));
            } catch (Exception ex) {
                try {
                    wb= new HSSFWorkbook(new POIFSFileSystem(new BufferedInputStream(new FileInputStream(file))));
                } catch (Exception e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            } //兼容读取设置结束
            Cell cell = null;
            System.out.println("本Excel总共有"+wb.getNumberOfSheets()+" 个Sheet 。");
            for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {//循环所有的sheet,一个excle中可能有多个sheet
                HashMap<String,String> map=new HashMap<String,String>();
                HashMap<String,String> tmap=new HashMap<String,String>();
                Sheet sheet = wb.getSheetAt(sheetIndex);
                int firstRowNum = sheet.getFirstRowNum();
                int lastRowNum = sheet.getLastRowNum();

                firstRowNum=firstRowNum>ignoreRows?firstRowNum:ignoreRows;
                Row row = null;
                for (int i = firstRowNum; i <= lastRowNum; i++) {
                    row = sheet.getRow(i);          //取得第i行
                    cell = row.getCell(column);        //取得i行的第column列
                    String value ="";//保存i行的第column列的值
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case HSSFCell.CELL_TYPE_STRING:
                                value = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    Date date = cell.getDateCellValue();
                                    if (date != null) {
                                        value = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss").format(date);
                                    } else {
                                        value = "";
                                    }
                                } else {
                                    value = new DecimalFormat("0").format(cell.getNumericCellValue());
                                }
                                break;
                            case HSSFCell.CELL_TYPE_FORMULA:
                                // 导入时如果为公式生成的数据则无值
                                if (!cell.getStringCellValue().equals("")) {
                                    value = cell.getStringCellValue();
                                } else {
                                    value = cell.getNumericCellValue() + "";
                                }
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                value = "";
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                value = (cell.getBooleanCellValue() == true ? "Y": "N");
                                break;
                            default:
                                value = "";
                        }

                    }
			        /*在excel中，计数是从0开始的，为了使结果与Excel中显示的行数保持一致，让行数newNum=为（i+1）
			         */
                    int newNum=i+1;
                    if(map.containsKey(value)){//如果Map集合中包含指定的键名，则返回true；否则返回false。
                        String lineNum=map.get(value);//拿到先前保存的行号
                        //System.out.println("先前保存的行号value="+value+" lineNum="+lineNum);
                        if(tmap.containsKey(value)){
                            String str=tmap.get(value);//拿到先前保存的所有行号记录
                            tmap.put(value, str+" ,"+newNum);//更新后，显示效果：——》行重复：在第 2 ，3 , 5
                        }else{
                            tmap.put(value, "重复：行数位于第  "+lineNum+" ,"+newNum);//最后显示效果：——》行重复：在第 2 ，3
                        }
                    }
                    map.put(value, newNum+"");//把i行的第column列的值与行号保存到map中
                }
                Iterator<Map.Entry<String, String>> it=tmap.entrySet().iterator();
                System.out.println("本Excel总共有"+wb.getNumberOfSheets()+" 个Sheet,第 "+(sheetIndex+1)+" 个Sheet中：");
                while(it.hasNext()){
                    Map.Entry<String, String> entry = (Map.Entry<String, String>) it.next();
                    System.out.println("字段："+entry.getKey()+" "+entry.getValue());
                }
                flag=true;
            }

            return flag;
        }
        return flag;
    }

    public static void main( String arg[]){
        CheckChongfuExcel eo=new CheckChongfuExcel();
        File file=new File("C:\\Users\\dell\\Desktop\\user1.xls");
        System.out.println("asfsaf  " +eo.judgeRepeatExcle(file, 0, 1));
    }

}
