package com.excel;

import com.excel.util.ExcelUtil;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

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

}
