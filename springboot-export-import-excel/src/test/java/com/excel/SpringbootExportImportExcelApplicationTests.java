package com.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

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

}
