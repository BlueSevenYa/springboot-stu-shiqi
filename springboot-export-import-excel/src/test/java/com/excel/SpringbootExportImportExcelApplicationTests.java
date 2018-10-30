package com.excel;

import com.excel.entity.User;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

@RunWith(SpringRunner.class)
@SpringBootTest
public class SpringbootExportImportExcelApplicationTests {

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
		System.out.println("ss");
		FileInputStream in=new FileInputStream("C:\\Users\\dell\\Desktop\\user1.xls");
		HSSFWorkbook workbook=new HSSFWorkbook(in);
		HSSFSheet sheet=workbook.getSheetAt(0);
		for(int i=1;i<=sheet.getLastRowNum();i++){
			Row row=sheet.getRow(i);
			User user=new User();
			user.setName((row.getCell(0)).toString());
			user.setAddress((row.getCell(1)).toString());
			System.out.println(user);
		}
	}

	@Test
	public void test3() throws IOException {
		FileInputStream in=new FileInputStream("C:\\Users\\dell\\Desktop\\user1.xls");
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
		in.close();
	}

}
