package service;

import java.io.FileOutputStream;
import java.io.IOException;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import model.Prograd;

//			Progression -1 
//Go to src/service. Open the ExcelGenerator and fill the logic inside the excelGenerate method.
//
//Stick to the instructions clearly. If you face any issue contact your mentor to get the guidance. 

public class ExcelGenerator {
	
	FileOutputStream out;
	public HSSFWorkbook excelGenerate(Prograd prograd, List<Prograd> list) throws IOException {
		try {


			
			HSSFWorkbook hwb = new HSSFWorkbook();
			
			HSSFSheet sheet = hwb.createSheet();
			HSSFRow hrow = sheet.createRow(0);
			hrow.createCell(0).setCellValue("Id");
			hrow.createCell(1).setCellValue("Name");
			hrow.createCell(2).setCellValue("Rate");
			hrow.createCell(3).setCellValue("Comment");
			hrow.createCell(4).setCellValue("Recommend");
			int i =0;
			for(Prograd pro:list) {
				HSSFRow dRow = sheet.createRow(++i);
				dRow.createCell(0).setCellValue(pro.getId());
				dRow.createCell(1).setCellValue(pro.getName());
				dRow.createCell(2).setCellValue(pro.getRate());
				dRow.createCell(3).setCellValue(pro.getComment());
				dRow.createCell(4).setCellValue(pro.getRecommend());
		
			}
			
			String filename = "E:/Information_Technology/progard.xls";
			// Do not modify the lines given below
			 out = new FileOutputStream(filename);
			hwb.write(out);
		
			return hwb;
			
		}
		catch (Exception e) {
				e.printStackTrace();
			}
		finally {
			out.close();
		}
		return null;
		
	}
}
