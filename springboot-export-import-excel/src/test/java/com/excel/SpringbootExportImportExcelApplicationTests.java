package com.excel;

import com.excel.util.ExcelUtil;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

@RunWith(SpringRunner.class)
@SpringBootTest
public class SpringbootExportImportExcelApplicationTests {

	private static final Logger log= LoggerFactory.getLogger(SpringbootExportImportExcelApplicationTests.class);

	@Test
	public void contextLoads() {
	}

	@Test
	public void test1() throws IOException {
		HSSFWorkbook workbook=new HSSFWorkbook();
		HSSFSheet sheet=workbook.createSheet("test");
		HSSFRow row=sheet.createRow(0);
		HSSFCell cell=row.createCell(0);
		cell.setCellValue("蓝十七");
		FileOutputStream output=new FileOutputStream("d:\\workbook.xls");
		workbook.write(output);
		output.flush();
	}

	@Test
	public void test2() throws IOException {
/*		System.out.println("ss");
		FileInputStream in=new FileInputStream("C:\\Users\\dell\\Desktop\\user1.xls");
		HSSFWorkbook workbook=new HSSFWorkbook(in);
		HSSFSheet sheet=workbook.getSheetAt(0);
		for(int i=1;i<=sheet.getLastRowNum();i++){
			Row row=sheet.getRow(i);
			User user=new User();
			user.setName((row.getCell(0)).toString());
			user.setAddress((row.getCell(1)).toString());
			System.out.println(user);
		}*/
	}

	@Test
	public void test3() throws IOException {
		/*FileInputStream in=new FileInputStream("C:\\Users\\dell\\Desktop\\user1.xls");
		HSSFWorkbook workbook=new HSSFWorkbook(in);
		HSSFSheet sheet=workbook.getSheetAt(0);
		for(Row r:sheet){
			System.out.println(r.getRowNum());
			if(r.getRowNum()<1){
				continue;
			}
			User user=new User();
			user.setName(r.getCell(0).getStringCellValue());
			user.setAddress(r.getCell(1).getStringCellValue());
			System.out.println(user);
		}
		in.close();*/
	}

	@Test
	public void test4() throws IOException {
		FileInputStream in=new FileInputStream("C:\\Users\\dell\\Desktop\\user1.xls");
		HSSFWorkbook workbook=new HSSFWorkbook(in);
		HSSFSheet sheet=workbook.getSheetAt(0);
		for(Row r:sheet){
			System.out.println(r.getRowNum());
			if(r.getRowNum()<1){
				continue;
			}
			System.out.println(r.getCell(0).getStringCellValue());
		}
		in.close();
	}

	@Test
	public void test5() throws IOException {
		InputStream in=new FileInputStream("C:\\Users\\dell\\Desktop\\user1.xls");
		Workbook workbook= ExcelUtil.getWorkBookByIo(in);
		Sheet sheet=workbook.getSheetAt(0);
		for(Row r:sheet){
			System.out.println(r.getRowNum());
			if(r.getRowNum()<1){
				continue;
			}
			System.out.println(r.getCell(0).getStringCellValue());
		}
		in.close();
	}

	@Test
	public void test6() throws IOException {
		InputStream in=new FileInputStream("C:\\Users\\dell\\Desktop\\template.xls");
		Workbook workbook= ExcelUtil.getWorkBookByIo(in);
		boolean s=ExcelUtil.checkTemplateRight(workbook,ExcelUtil.notHeadExcelNum,ExcelUtil.headName);
		log.info(s+" ");
		in.close();
	}

	@Test
	public void test7() throws IOException {
		//创建工作簿对象
		HSSFWorkbook wb=new HSSFWorkbook();
		//创建工作表对象
		HSSFSheet sheet=wb.createSheet("我的工作表");
		//创建绘图对象
		HSSFPatriarch p=sheet.createDrawingPatriarch();
		//设置样式-颜色
		HSSFCellStyle style = wb.createCellStyle();/*
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);*/
		style.setFillForegroundColor(HSSFColor.YELLOW.index);
		//创建单元格对象,批注插入到4行,1列,B5单元格
		HSSFCell cell=sheet.createRow(4).createCell(1);
		cell.setCellStyle(style);
		//插入单元格内容
		cell.setCellValue(new HSSFRichTextString("批注"));
		//获取批注对象
		//(int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2)
		//前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
		HSSFComment comment=p.createComment(new HSSFClientAnchor(0,0,0,0,(short)3,3,(short)5,6));
		//输入批注信息
		comment.setString(new HSSFRichTextString("插件批注成功!插件批注成功!"));
		//添加作者,选中B5单元格,看状态栏
		comment.setAuthor("admin");
		comment.setFillColor(244,244,88);
		//将批注添加到单元格对象中
		cell.setCellComment(comment);
		//创建输出流
		FileOutputStream out=new FileOutputStream("C:\\Users\\dell\\Desktop\\writerPostil.xls");

		wb.write(out);
		//关闭流对象
		out.close();
	}


	@Test
	public void test8() throws IOException {
		//创建工作簿对象
		HSSFWorkbook wb=new HSSFWorkbook();
		//创建工作表对象
		HSSFSheet sheet=wb.createSheet("我的工作表");
		//创建绘图对象
		HSSFPatriarch p=sheet.createDrawingPatriarch();
		//设置样式-颜色
		CellStyle style = ExcelUtil.setCellStyle(wb);
		//创建单元格对象,批注插入到4行,1列,B5单元格
		HSSFCell cell=sheet.createRow(4).createCell(1);
		cell.setCellStyle(style);
		//插入单元格内容
		cell.setCellValue(new HSSFRichTextString("批注"));
		HSSFComment comment=ExcelUtil.setCellComment(p,"测试");
		//将批注添加到单元格对象中
		cell.setCellComment(comment);
		//创建输出流
		FileOutputStream out=new FileOutputStream("C:\\Users\\dell\\Desktop\\writerPostil.xls");

		wb.write(out);
		//关闭流对象
		out.close();
	}

	@Test
	public void test9(){
		String s="000a";
		if(s.matches(ExcelUtil.regUserName)){
			log.info("true");
		}else{
			log.info("false");
		}
	}

	@Test
	public void test12(){
		FileSystemView fsv = FileSystemView.getFileSystemView();
		File com=fsv.getHomeDirectory();    //这便是读取桌面路径的方法了
		System.out.println(com.getPath());
	}











	@Test
	public void test11(){
		try{
			String urlFile="http://img.zcool.cn/community/01d881579dc3620000018c1b430c4b.JPG@3000w_1l_2o_100sh.jpg";
			String saveName="1.jpg";
			String urlLoad="C:\\Users\\dell\\Desktop\\";
			downLoadFromUrl(urlFile,saveName,urlLoad);
		}catch (Exception e) {
			// TODO: handle exception
		}
	}

	/**
	 * 从网络Url中下载文件
	 * @param urlStr
	 * @param fileName
	 * @param savePath
	 * @throws IOException
	 */
	public static void  downLoadFromUrl(String urlStr,String fileName,String savePath) throws IOException {
		URL url = new URL(urlStr);
		HttpURLConnection conn = (HttpURLConnection)url.openConnection();
		//设置超时间为3秒
		conn.setConnectTimeout(3*1000);
		//防止屏蔽程序抓取而返回403错误
		conn.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)");

		//得到输入流
		InputStream inputStream = conn.getInputStream();
		//获取自己数组
		byte[] getData = readInputStream(inputStream);

		//文件保存位置
		File saveDir = new File(savePath);
		if(!saveDir.exists()){
			saveDir.mkdir();
		}
		File file = new File(saveDir+ File.separator+fileName);
		FileOutputStream fos = new FileOutputStream(file);
		fos.write(getData);
		if(fos!=null){
			fos.close();
		}
		if(inputStream!=null){
			inputStream.close();
		}


		System.out.println("info:"+url+" download success");

	}



	/**
	 * 从输入流中获取字节数组
	 * @param inputStream
	 * @return
	 * @throws IOException
	 */
	public static  byte[] readInputStream(InputStream inputStream) throws IOException {
		byte[] buffer = new byte[1024];
		int len = 0;
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		while((len = inputStream.read(buffer)) != -1) {
			bos.write(buffer, 0, len);
		}
		bos.close();
		return bos.toByteArray();
	}



	@Test
	public void test10(){
		File file=new File("C:\\Users\\dell\\Desktop\\user1.xls");
		System.out.println("asfsaf  " +judgeRepeatExcle(file, 0, 1));
	}


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

}
