package test.excel;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.lunjar.excelutil.ReadUtil;
import com.lunjar.excelutil.WriteUtil;

public class CommTest {

	public static void main(String[] args) throws Exception {
		
//		writeTest();
		
	}
	
	public static void writeTest(){
		try{
			String fileName = "e://通讯录备份.xls";
			String orifile = "e://通讯录备份.txt";
			
			Pattern p1 = Pattern.compile("(?<=姓名:\\s)\\S+(?=---)");
			Pattern p2 = Pattern.compile("(?<=电话:\\s)\\d+");
			
			List<String[]> contact = new ArrayList<String[]>();
			Map<String, List<String>> map = new HashMap<String, List<String>>();
			
			BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(orifile)));
			String line = null;
			line = br.readLine();
			while(null != line){
				String name = "";
				String phone = "";
				
				Matcher m1 = p1.matcher(line);
				if(m1.find()){
					name = m1.group();
				}
				Matcher m2 = p2.matcher(line);
				if(m2.find()){
					phone = m2.group();
				}
				System.out.println("(name : "+name + ")(phone : "+phone+")");
				System.out.println();
				
				if(null != name && name.length() > 0){
					List<String> strArr = map.get(name);
					if(null == strArr){
						strArr = new ArrayList<String>();
					}
					strArr.add(phone);
					map.put(name, strArr);
				}
				
				line = br.readLine();
			}
			
			for(String key : map.keySet()){
				List<String> list = map.get(key);
				if(null != list){
					int i = 0;
					for(String phone : list){
						String[] arr = new String[2];
						arr[0] = "";
						arr[1] = phone;
						if(i == 0){
							arr[0] = key;
						}
						contact.add(arr);
						i++;
					}
					String[] tarr = new String[2];
					tarr[0] = "";
					tarr[1] = "";
					contact.add(tarr);
				}
				
			}
			
			
			boolean bool = WriteUtil.writeXLS(fileName, "通讯录备份", contact);
			System.out.println(bool);
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void readTest(){
		try {
			List<String> list = ReadUtil.readXLS("e://123.xls",0);
			if(null != list){
				for(String s : list){
					System.out.println(s);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}
	
}
