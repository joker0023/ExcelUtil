package test.excel;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.lunjar.excelutil.ReadUtil;
import com.lunjar.excelutil.WriteUtil;

public class OtherTest {

	public static void main(String[] args) {
		try{
//			appendTest();
//			excelTest();
//			writeTxtTest();
			
			String oriFile = "C://Documents and Settings/Administrator/桌面/欺诈历史匹配数据.xls";
			
			String target = "C://Documents and Settings/Administrator/桌面/target.txt";
			
			writeTxtTest(oriFile, target, 5);
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 根据源文件生成多份文件
	 * @param oriFile 源文件
	 * @param target	目标文件
	 * @param times		翻几倍
	 * @throws Exception
	 */
	public static void writeTxtTest(String oriFile, String target, int times) throws Exception {
		BufferedWriter bw = null;
		try{
			String[] randomArr = {"北京", "长沙市", "福州", "广州", "广州市", "哈尔滨",
					"黄石", "惠州", "揭阳", "江门", "南京", "南通", "三门峡", "山东",
					"汕头", "上海", "深圳", "沈阳", "苏州", "盐城", "云浮", "中山", "珠海"};
			
			//读取
			int[] arr = new int[72];
			for(int i = 0; i < 72; i++){
				arr[i] = i;
			}
			List<String[]> list = ReadUtil.readXLS(oriFile, arr);
			
			//写入
			bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(target)));
			String line = null;
			String[] row = null;
			int size = 0;
			
			row = list.get(2);
			size = row.length;
			line = "";
			for(int i = 0; i < size; i++){
				line += row[i] + " | ";
			}
			bw.write(line + "\r\n\r\n");
			
			int rowLength = list.size();
			BigInteger addNum = null;
			for(int i = 0; i < times; i++){
				System.out.println("翻第" + (i + 1) + "倍....");
				addNum = new BigInteger(Integer.toString(1));
				for(int j = 3; j < rowLength; j++){
					row = list.get(j);
					if(StringUtils.isNotBlank(row[0])){
						row[0] = new BigInteger(row[0]).add(addNum).toString();
					}
					if(StringUtils.isNotBlank(row[1])){
						row[1] = new BigInteger(row[1]).add(addNum).toString();
					}
					row[2] = randomArr[((int)Math.round(Math.random() * 10)) % randomArr.length];
					row[3] = randomArr[((int)Math.round(Math.random() * 10)) % randomArr.length];
					if(StringUtils.isNotBlank(row[8])){
						row[8] = new BigInteger(row[8]).add(addNum).toString();
					}
					if(StringUtils.isNotBlank(row[9])){
						row[9] = new BigInteger(row[9]).add(addNum).toString();
					}
					
					size = row.length;
					line = "";
					for(int k = 0; k < size; k++){
						line += row[k] + " | ";
					}
					bw.write(line + "\r\n");
				}
			}
			bw.flush();
			
			System.out.println("执行完成!!!");
		}catch (Exception e) {
			e.printStackTrace();
		}finally {
			if(null != bw){
				bw.close();
			}
		}
	}
	
	public static void writeTxtTest() throws Exception {
		try{
			String[] randomArr = {"广东", "江苏"};
			
			String oriFile = "C://Documents and Settings/Administrator/桌面/欺诈历史匹配数据.xls";
			
			String target = "C://Documents and Settings/Administrator/桌面/target.txt";
			
			//读取
			int[] arr = new int[72];
			for(int i = 0; i < 72; i++){
				arr[i] = i;
			}
			List<String[]> list = ReadUtil.readXLS(oriFile, arr);
			
			//写入
			BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(target)));
			String line = null;
			String[] row = null;
			int size = 0;
			
			row = list.get(2);
			size = row.length;
			line = "";
			for(int i = 0; i < size; i++){
				line += row[i] + " | ";
			}
			bw.write(line + "\r\n");
			
			int rowLength = list.size();
			int rowIndex = 1;
			BigInteger addNum = null;
			for(int i = 0; i < 100; i++){
				addNum = new BigInteger(Integer.toString(1));
				for(int j = 3; j < rowLength; j++){
					System.out.println("第 " + (++rowIndex) + " 行");
					row = list.get(j);
					if(StringUtils.isNotBlank(row[0])){
						row[0] = new BigInteger(row[0]).add(addNum).toString();
					}
					if(StringUtils.isNotBlank(row[1])){
						row[1] = new BigInteger(row[1]).add(addNum).toString();
					}
					row[2] = randomArr[((int)Math.round(Math.random() * 10)) % 2];
					row[3] = randomArr[((int)Math.round(Math.random() * 10)) % 2];
					if(StringUtils.isNotBlank(row[8])){
						row[8] = new BigInteger(row[8]).add(addNum).toString();
					}
					if(StringUtils.isNotBlank(row[9])){
						row[9] = new BigInteger(row[9]).add(addNum).toString();
					}
					
					size = row.length;
					line = "";
					for(int k = 0; k < size; k++){
						line += row[k] + " | ";
					}
					bw.write(line + "\r\n");
				}
			}
			bw.flush();
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void appendTest() throws IOException {
		FileOutputStream out = null;
		try{
			String[] randomArr = {"广东", "江苏"};
			
			String oriFile = "C://Documents and Settings/Administrator/桌面/欺诈历史匹配数据.xls";
			
			String target = "C://Documents and Settings/Administrator/桌面/target.xls";
			
			//读取
			int[] arr = new int[72];
			for(int i = 0; i < 72; i++){
				arr[i] = i;
			}
			List<String[]> list = ReadUtil.readXLS(oriFile, arr);
			
			//写入
			XSSFWorkbook workbook = new XSSFWorkbook();
			
			XSSFSheet sheet =  workbook.createSheet();
			
			addRow(sheet, list.get(0), 0);
			addRow(sheet, list.get(1), 1);
			addRow(sheet, list.get(2), 2);
			
			int rowLength = list.size();
			int rowIndex = 3;
			String[] row = null;
			BigInteger addNum = null;
			for(int i = 0; i < 1; i++){
				addNum = new BigInteger(Integer.toString(i + 1));
				for(int j = 3; j < rowLength - 500; j++){
					System.out.println("第 " + (rowIndex + 1) + " 行");
					row = list.get(j);
					if(StringUtils.isNotBlank(row[0])){
						row[0] = new BigInteger(row[0]).add(addNum).toString();
					}
					if(StringUtils.isNotBlank(row[1])){
						row[1] = new BigInteger(row[1]).add(addNum).toString();
					}
					row[2] = randomArr[((int)Math.round(Math.random() * 10)) % 2];
					row[3] = randomArr[((int)Math.round(Math.random() * 10)) % 2];
					if(StringUtils.isNotBlank(row[8])){
						row[8] = new BigInteger(row[8]).add(addNum).toString();
					}
					if(StringUtils.isNotBlank(row[9])){
						row[9] = new BigInteger(row[9]).add(addNum).toString();
					}
					
					addRow(sheet, row, rowIndex++);
				}
			}
			
			out = new FileOutputStream(target);
			workbook.write(out);
		}catch (Exception e) {
			e.printStackTrace();
		}finally {
			out.close();
		}
	}
	
	/**
	 * 写入一行
	 * @param sheet
	 * @param rowContent
	 * @param index
	 * @throws Exception
	 */
	public static void addRow(XSSFSheet sheet, String[] rowContent, int index) throws Exception {
		if(null == rowContent){
			return;
		}
		XSSFRow row = sheet.createRow(index);
		
		int cellNum = rowContent.length;
		String cellContent = null;
		XSSFCell cell = null;
		for(int i = 0; i < cellNum; i++){
			cellContent = rowContent[i];
			cell = row.createCell(i);
			
			if(null != cellContent){
				cell.setCellValue(cellContent);
			}
		}
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	public static void excelTest() throws IOException {
		try{
			String oriFile = "C://Documents and Settings/Administrator/桌面/欺诈历史匹配数据.xls";
			
			String target = "C://Documents and Settings/Administrator/桌面/target.xls";
			
			int[] arr = new int[72];
			for(int i = 0; i < 72; i++){
				arr[i] = i;
			}
			List<String[]> list = ReadUtil.readXLS(oriFile, arr);
			
			List<String[]> newList = new ArrayList<String[]>();
			newList.add(list.get(0));
			newList.add(list.get(1));
			newList.add(list.get(2));
			
			String[] randomArr = {"广东", "江苏"};
			String[] row = null;
			
			BigInteger addNum = null;
			for(int i = 0; i < 10; i++){
				addNum = new BigInteger(Integer.toString(i + 1));
				for(int j = 3; j < 1003; j++){
					System.out.println("第 " + (j + 1) + " 行");
					row = list.get(j);
					if(StringUtils.isNotBlank(row[0])){
						row[0] = new BigInteger(row[0]).add(addNum).toString();
					}
					if(StringUtils.isNotBlank(row[1])){
						row[1] = new BigInteger(row[1]).add(addNum).toString();
					}
					row[2] = randomArr[((int)Math.round(Math.random() * 10)) % 2];
					row[3] = randomArr[((int)Math.round(Math.random() * 10)) % 2];
					if(StringUtils.isNotBlank(row[8])){
						row[8] = new BigInteger(row[8]).add(addNum).toString();
					}
					if(StringUtils.isNotBlank(row[9])){
						row[9] = new BigInteger(row[9]).add(addNum).toString();
					}
					
					newList.add(row);
				}
			}
			
			WriteUtil.writeXLS(target, "CAPS.GDB_HISMATCHAPP", newList);
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
}
