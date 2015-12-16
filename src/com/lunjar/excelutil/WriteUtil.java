package com.lunjar.excelutil;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteUtil {
	
	/**
	 * 写入xlsx文件
	 * @param filePath	文件名(完整路径)
	 * @param sheetName	sheet的名称
	 * @param sheetContent	sheet的内容
	 * @return
	 * @throws Exception
	 */
	public static boolean writeXLSX(String filePath,String sheetName,List<String[]> sheetContent) throws Exception{
		Map<String,List<String[]>> sheetMap = new HashMap<String, List<String[]>>();
		sheetMap.put(sheetName, sheetContent);
		return writeXSSF(filePath, sheetMap);
	}
	
	/**
	 * 写入xlsx文件
	 * @param out	输入流
	 * @param sheetName	sheet的名称
	 * @param sheetContent	sheet的内容
	 * @return
	 * @throws Exception
	 */
	public static boolean writeXLSX(OutputStream out,String sheetName,List<String[]> sheetContent) throws Exception{
		Map<String,List<String[]>> sheetMap = new HashMap<String, List<String[]>>();
		sheetMap.put(sheetName, sheetContent);
		return writeXSSF(out, sheetMap);
	}
	
	/**
	 * 写入xlsx文件
	 * @param filePath	文件名(完整路径)
	 * @param sheetMap	key--sheet名称 	value(list)--sheet内容	string[]--一行的内容
	 * @return
	 * @throws Exception
	 */
	public static boolean writeXLSX(String filePath, Map<String,List<String[]>> sheetMap) throws Exception{
		return writeXSSF(filePath, sheetMap);
	}
	
	/**
	 * 写入xlsx文件
	 * @param out	输入流
	 * @param sheetMap	key--sheet名称 	value(list)--sheet内容	string[]--一行的内容
	 * @return
	 * @throws Exception
	 */
	public static boolean writeXLSX(OutputStream out, Map<String,List<String[]>> sheetMap) throws Exception{
		return writeXSSF(out, sheetMap);
	}

	/**
	 * 写入xlsx文件(sheetMap为空不创建文件)
	 * @param filePath	文件名(完整路径)
	 * @param sheetMap	key--sheet名称 	value(list)--sheet内容	string[]--一行的内容
	 * @return
	 * @throws Exception
	 */
	private static boolean writeXSSF(String filePath, Map<String,List<String[]>> sheetMap) throws Exception{
		FileOutputStream out  =null;
		out = new FileOutputStream(filePath);
		return writeXSSF(out, sheetMap);
	}
	
	/**
	 * 写入xlsx文件(sheetMap为空不创建文件)
	 * @param out	输入流
	 * @param sheetMap	key--sheet名称 	value(list)--sheet内容	string[]--一行的内容
	 * @return
	 * @throws Exception
	 */
	private static boolean writeXSSF(OutputStream out, Map<String,List<String[]>> sheetMap) throws Exception{
		try {
			if(null != sheetMap){
				//创建Excel工作簿
				XSSFWorkbook workbook = new XSSFWorkbook(); 
				
				//遍历每一个sheet
				for(String sheetName: sheetMap.keySet()){
					//得到一个将要写入的sheet内容
					List<String[]> sheetContent = sheetMap.get(sheetName);
					//新建一个sheet
					XSSFSheet sheet =  workbook.createSheet(sheetName);
					
					if(null != sheetContent){
						//遍历每一行
						int rowNum = sheetContent.size();
						for(int i = 0; i < rowNum; i++){
							//得到一 行将要写入的内容
							String[] rowContent  = sheetContent.get(i);
							//新建一行
							XSSFRow row = sheet.createRow(i);
							
							if(null != rowContent){
								//遍历每一单元格
								int cellNum = rowContent.length;
								for(int j = 0; j < cellNum; j++){
									String cellContent = rowContent[j];
									XSSFCell cell = row.createCell(j);
									
									if(null != cellContent){
										cell.setCellValue(cellContent);
									}
								}
							}
						}
					}
				}
				//写入文件
				workbook.write(out);
			}
		}catch (Exception e) {
			throw new Exception(e);
		}finally {
			if (out != null) {
				try {
					out.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
		
		return true;
	}
	
	/**
	 * 写入xls文件
	 * @param filePath	文件名(完整路径)
	 * @param sheetName	sheet的名称
	 * @param sheetContent	sheet的内容
	 * @return
	 * @throws Exception
	 */
	public static boolean writeXLS(String filePath, String sheetName, List<String[]> sheetContent) throws Exception{
		Map<String,List<String[]>> sheetMap = new HashMap<String, List<String[]>>();
		sheetMap.put(sheetName, sheetContent);
		return writeHSSF(filePath, sheetMap);
	}
	
	/**
	 * 写入xls文件
	 * @param out	输入流
	 * @param sheetName	sheet的名称
	 * @param sheetContent	sheet的内容
	 * @return
	 * @throws Exception
	 */
	public static boolean writeXLS(OutputStream out, String sheetName, List<String[]> sheetContent) throws Exception{
		Map<String,List<String[]>> sheetMap = new HashMap<String, List<String[]>>();
		sheetMap.put(sheetName, sheetContent);
		return writeHSSF(out, sheetMap);
	}
	
	
	/**
	 * 写入xls文件(sheetMap为空不创建文件)
	 * @param filePath	文件名(完整路径)
	 * @param sheetMap	key--sheet名称 	value(list)--sheet内容	string[]--一行的内容
	 * @return
	 * @throws Exception
	 */
	public static boolean writeXLS(String filePath, Map<String,List<String[]>> sheetMap) throws Exception{
		return writeHSSF(filePath, sheetMap);
	}
	
	/**
	 * 写入xls文件(sheetMap为空不创建文件)
	 * @param out	输入流
	 * @param sheetMap	key--sheet名称 	value(list)--sheet内容	string[]--一行的内容
	 * @return
	 * @throws Exception
	 */
	public static boolean writeXLS(OutputStream out, Map<String,List<String[]>> sheetMap) throws Exception{
		return writeHSSF(out, sheetMap);
	}
	
	/**
	 * 写入xls文件(sheetMap为空不创建文件)
	 * @param filePath	文件名(完整路径)
	 * @param sheetMap	key--sheet名称 	value(list)--sheet内容	string[]--一行的内容
	 * @return
	 * @throws Exception
	 */
	private static boolean writeHSSF(String filePath, Map<String,List<String[]>> sheetMap) throws Exception{
		FileOutputStream out  =null;
		out = new FileOutputStream(filePath);
		return writeHSSF(out, sheetMap);
	}
	
	/**
	 * 写入xls文件(sheetMap为空不创建文件)
	 * @param filePath	输入流
	 * @param sheetMap	key--sheet名称 	value(list)--sheet内容	string[]--一行的内容
	 * @return
	 * @throws Exception
	 */
	private static boolean writeHSSF(OutputStream out, Map<String,List<String[]>> sheetMap) throws Exception{
		try {
			if(null != sheetMap){
				//创建Excel工作簿
				HSSFWorkbook workbook = new HSSFWorkbook(); 
				
				//遍历每一个sheet
				for(String sheetName: sheetMap.keySet()){
					//得到一个将要写入的sheet内容
					List<String[]> sheetContent = sheetMap.get(sheetName);
					//新建一个sheet
					HSSFSheet sheet =  workbook.createSheet(sheetName);
					
					if(null != sheetContent){
						//遍历每一行
						int rowNum = sheetContent.size();
						for(int i = 0; i < rowNum; i++){
							//得到一 行将要写入的内容
							String[] rowContent  = sheetContent.get(i);
							//新建一行
							HSSFRow row = sheet.createRow(i);
							
							if(null != rowContent){
								//遍历每一单元格
								int cellNum = rowContent.length;
								for(int j = 0; j < cellNum; j++){
									String cellContent = rowContent[j];
									HSSFCell cell = row.createCell(j);
									
									if(null != cellContent){
										cell.setCellValue(cellContent);
									}
								}
							}
						}
					}
				}
				//写入文件
				
				workbook.write(out);
			}
		}catch (Exception e) {
			throw new Exception(e);
		}finally {
			if (out != null) {
				try {
					out.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
		
		return true;
	}
}
