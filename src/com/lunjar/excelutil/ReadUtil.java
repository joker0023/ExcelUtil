package com.lunjar.excelutil;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadUtil {

	/**
	 * 读取xlsx文件(每一个sheet的每一行的第一列)
	 * @param filePath(完整路径)
	 * @return
	 * @throws Exception 
	 */
	public static List<String> readXLSX(String filePath) throws Exception{
		return readXLSX(filePath, 0);
	}
	
	/**
	 * 读取xlsx文件(每一个sheet的每一行的第columnNum列)
	 * @param filePath	文件名(完整路径)
	 * @param columnNum	第几列 (0开始)
	 * @return
	 * @throws Exception 
	 */
	public static List<String> readXLSX(String filePath, int columnNum) throws Exception{
		List<String[]> list = readXSSF(filePath, columnNum);
		
		List<String> rtnList = new ArrayList<String>();
		if(null != list){
			for(String[] strArr : list){
				rtnList.add(strArr[0]);
			}
		}
		
		return rtnList;
	}
	
	/**
	 * 读取xlsx文件(每一个sheet的每一行的第columnNum...列,不传默认读第一列)
	 * @param filePath(完整路径)
	 * @param columnNum 第几列(从0开始)
	 * @return
	 * @throws Exception 
	 */
	public static List<String[]> readXLSX(String filePath, int...columnNum) throws Exception{
		return readXSSF(filePath, columnNum);
	}
	
	/**
	 * 读取xlsx文件(每一个sheet的每一行的第columnNum...列,不传默认读第一列)
	 * @param filePath(完整路径)
	 * @param columnNum 第几列(从0开始)
	 * @return
	 * @throws FileNotFoundException 
	 */
	private static List<String[]> readXSSF(String filePath, int...columnNum) throws Exception{
		int columnNums = 1;
		if(null != columnNum){
			columnNums = columnNum.length;
		}else{
			columnNum = new int[]{0};
		}
		List<String[]> list = new ArrayList<String[]>();
		
		FileInputStream in  =null;
		try {
			in = new FileInputStream(filePath);
			//创建对Excel工作簿文件的引用 
			XSSFWorkbook workbook = new XSSFWorkbook(in); 
			
			for(int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
				//获得一个sheet
				XSSFSheet sheet = workbook.getSheetAt(sheetNum); 
				if(null != sheet){
					for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
						//获得一行
						XSSFRow row = sheet.getRow(rowNum); 
		                if (null != row) {
		                	String[] content = new String[columnNums];
		                	for(int i = 0; i < columnNums; i++){
		                		StringBuffer sbf = new StringBuffer();
		                		int colNum = columnNum[i];
		                		//获取一列
		                    	XSSFCell cell = row.getCell(colNum);	
		                        if (null != cell) {
		                            if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
		                            	sbf.append(cell.getNumericCellValue());
		                            }else if(cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN){
		                            	sbf.append(cell.getBooleanCellValue());
		                            }else {
		                            	sbf.append(cell.getStringCellValue());
		                            }
		                            content[i] = sbf.toString();
		                        }
		                	}
		                	list.add(content);
		                }
					}
				}
			}
			in.close();
		} catch (FileNotFoundException e) {
			throw new Exception(e);
		} catch (IOException e) {
			throw new Exception(e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
		return list;
	}
	
	/**
	 * 读取xls文件(每一个sheet的每一行的第一列)
	 * @param filePath(完整路径)
	 * @return
	 * @throws Exception 
	 */
	public static List<String> readXLS(String filePath) throws Exception{
		return readXLS(filePath, 0);
	}
	
	/**
	 * 读取xls文件(每一个sheet的每一行的第columnNum列)
	 * @param filePath	文件名(完整路径)
	 * @param columnNum	第几列(0开始)
	 * @return
	 * @throws Exception 
	 */
	public static List<String> readXLS(String filePath, int columnNum) throws Exception{
		List<String[]> list = readHSSF(filePath, columnNum);
		
		List<String> rtnList = new ArrayList<String>();
		if(null != list){
			for(String[] strArr : list){
				rtnList.add(strArr[0]);
			}
		}
		
		return rtnList;
	}
	
	/**
	 * 读取xls文件(每一个sheet的每一行的第columnNum...列,不传默认读第一列)
	 * @param filePath(完整路径)
	 * @param columnNum 第几列(从0开始)
	 * @return
	 * @throws Exception 
	 */
	public static List<String[]> readXLS(String filePath, int...columnNum) throws Exception{
		return readHSSF(filePath, columnNum);
	}
	
	/**
	 * 读取xls文件(每一个sheet的每一行的第columnNum...列,不传默认读第一列)
	 * @param filePath(完整路径)
	 * @param columnNum 第几列(从0开始)
	 * @return
	 * @throws Exception 
	 */
	private static List<String[]> readHSSF(String filePath, int...columnNum) throws Exception{
		int columnNums = 1;
		if(null != columnNum){
			columnNums = columnNum.length;
		}else{
			columnNum = new int[]{0};
		}
		List<String[]> list = new ArrayList<String[]>();
		
		FileInputStream in  =null;
		try {
			in = new FileInputStream(filePath);
			//创建对Excel工作簿文件的引用 
			HSSFWorkbook workbook = new HSSFWorkbook(in); 
			
			for(int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++){
				//获得一个sheet
				HSSFSheet sheet = workbook.getSheetAt(sheetNum); 
				if(null != sheet){
					for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
						//获得一行
						HSSFRow row = sheet.getRow(rowNum); 
		                if (null != row) {
		                	String[] content = new String[columnNums];
		                	for(int i = 0; i < columnNums; i++){
		                		StringBuffer sbf = new StringBuffer();
		                		int colNum = columnNum[i];
		                		//获取一列
		                    	HSSFCell cell = row.getCell(colNum);
		                        if (null != cell) {
		                            if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
		                            	sbf.append(cell.getNumericCellValue());
		                            }else if(cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN){
		                            	sbf.append(cell.getBooleanCellValue());
		                            }else if(cell.getCellType() == HSSFCell.CELL_TYPE_ERROR)
		                            	sbf.append(cell.getErrorCellValue());
		                            else {
		                            	sbf.append(cell.getStringCellValue());
		                            }
		                            content[i] = sbf.toString();
		                        }
		                	}
		                	list.add(content);
		                }
					}
				}
			}
			in.close();
		} catch (FileNotFoundException e) {
			throw new Exception(e);
		} catch (IOException e) {
			throw new Exception(e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
		return list;
	}
	
}
