package com.din.comp;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;


/**
 * Author DINESH_11_nov_2016
 * 
 * */
/*
0 Approval Status 	
1 License Conflict 	
2 Component 	
3 Version 	
4 Home Page 	
5 Component Comment 	
6 License 	
7 External IDs 	
8 Usage 	
9 Ship Status 	
10 # Manual Code Match*/

public class BOMDBCreationq {

	public static void main(String[] args) throws IOException {
		File BOMFile;
		Scanner scan = new Scanner(System.in);
		System.out.println("Drag and drop folder which contains only html report:");
		String folderinput = scan.next();//

		
		File files[] = new File(folderinput).listFiles(new FilenameFilter() { public boolean accept(File dir, String name) { return name.endsWith(".html"); }});

		if(files.length != 0){
			for (int i = 0; i < files.length; i++) {
				System.out.println(files[i]);
			}
		}
		else System.out.println("no files found");

		String outputFilePath = folderinput+"/Output.xlsx";
		System.out.println("Check the Output.txt for result");

		ArrayList<String> as = new ArrayList<String>();//*
		ArrayList<String> lc = new ArrayList<String>();
		ArrayList<String> nm = new ArrayList<String>();//*
		ArrayList<String> v = new ArrayList<String>();//*
		ArrayList<String> hp = new ArrayList<String>();//*
		ArrayList<String> cc = new ArrayList<String>();//*
		ArrayList<String> l = new ArrayList<String>();//*
		ArrayList<String> ei = new ArrayList<String>();
		ArrayList<String> u = new ArrayList<String>();//*
		ArrayList<String> ss = new ArrayList<String>();//*
		ArrayList<String> mcm = new ArrayList<String>();

		FileOutputStream outputStream = new FileOutputStream(outputFilePath);
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Master BOM");

		int rowCount = 0, columnCount = 0;
		ArrayList<String> head = new ArrayList<String>();
		head.add("Approval Status");
		head.add("Component");
		head.add("Version");
		head.add("Home Page");
		head.add("Component Comment");
		head.add("License");
		head.add("Usage");
		head.add("Ship Status");
		head.add("File name");
		org.apache.poi.ss.usermodel.Row row = null;
		row = sheet.createRow(rowCount++);

		//
		CellStyle stylehead = workbook.createCellStyle();
		CellStyle stylewrap = workbook.createCellStyle();
		stylewrap.setWrapText(true); 

		stylehead.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		stylehead.setFillPattern(CellStyle.SOLID_FOREGROUND);
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		stylehead.setFont(font);

		Cell cell = null;
		for (String h : head) {
			cell = row.createCell(columnCount++);
			cell.setCellValue((String) h);
			cell.setCellStyle(stylehead);
		}//int k=0,l1=0,m=0,n1=0,o=0,p=0,q=0,r=0,s=0;



		Elements c_as = null;//*
		Elements c_lc = null;
		Elements c_nm = null;//*
		Elements c_v = null;//*
		Elements c_hp = null;//*
		Elements c_cc = null;//*
		Elements c_l = null;
		//Elements c_ei = null;
		Elements c_u = null;
		Elements c_ss = null;//*
		Elements c_mcm = null;
		String filename = null;

		//version
		Elements elements;
		//String bomversiontext;
		//version
		String text = null;

		org.jsoup.nodes.Document document;
		int temp = 1;
		for (int i = 0; i < files.length; i++) {
			as.clear();lc.clear();nm.clear();v.clear();hp.clear();cc.clear();l.clear();ei.clear();u.clear();ss.clear();mcm.clear();
			BOMFile = new File("/"+files[i]);

			System.out.println("Processing file"+files[i]);
			filename = (String)files[i].getName();
			document =  Jsoup.parse(BOMFile, "UTF-8", "");


			//version ex id check
			elements = document.select(".bomTable th:eq(7)");
			if (elements.size() == 0) {
				/*
				elements = document.select("table.reportTable tbody tr:nth-child(8) td:nth-child(2)"); 
				if (elements.size() ==0) {
					bomversiontext = "none";
				} else {
					bomversiontext = elements.get(1).text();
				}*/
				text = "NA";
				} else {
				//				 bomversiontext = elements.get(1).text();
				text = elements.get(0).text();
			}
			

			if (text.equalsIgnoreCase("External IDs")) {
				c_as = document.select(".bomTable td:eq(0)");//*
				c_lc = document.select(".bomTable td:eq(1)");
				c_nm = document.select(".bomTable td:eq(2)");//*
				c_v = document.select(".bomTable td:eq(3)");//*
				c_hp = document.select(".bomTable td:eq(4)");//*
				c_cc = document.select(".bomTable td:eq(5)");//*
				c_l = document.select(".bomTable td:eq(6)");//*
				//c_ei = document.select(".bomTable td:eq(7)");
				c_u = document.select(".bomTable td:eq(8)");//*
				c_ss = document.select(".bomTable td:eq(9)");//*
				c_mcm = document.select(".bomTable td:eq(10)");
			} else {
				c_as = document.select(".bomTable td:eq(0)");//*
				c_lc = document.select(".bomTable td:eq(1)");
				c_nm = document.select(".bomTable td:eq(2)");//*
				c_v = document.select(".bomTable td:eq(3)");//*
				c_hp = document.select(".bomTable td:eq(4)");//*
				c_cc = document.select(".bomTable td:eq(5)");//*
				c_l = document.select(".bomTable td:eq(6)");//*
				//external id not available
				c_u = document.select(".bomTable td:eq(7)");//*
				c_ss = document.select(".bomTable td:eq(8)");//*
				c_mcm = document.select(".bomTable td:eq(9)");
			}

			//version  ext id check






			/*
			 c_as = document.select(".bomTable td:eq(0)");//*
			 c_lc = document.select(".bomTable td:eq(1)");
			 c_nm = document.select(".bomTable td:eq(2)");//*
			 c_v = document.select(".bomTable td:eq(3)");//*
			 c_hp = document.select(".bomTable td:eq(4)");//*
			 c_cc = document.select(".bomTable td:eq(5)");//*
			 c_l = document.select(".bomTable td:eq(6)");//*
			 c_ei = document.select(".bomTable td:eq(7)");
			 c_u = document.select(".bomTable td:eq(8)");//*
			 c_ss = document.select(".bomTable td:eq(9)");//*
			 c_mcm = document.select(".bomTable td:eq(10)");*/





			for (Element nextTurn : c_as ) {
				//System.out.println(nextTurn.text());
				as.add(nextTurn.text());                           //list1
			}

			for (Element nextTurn : c_nm ) {
				//	System.out.println(nextTurn.text());
				nm.add(nextTurn.text());                           //list1
			}
			for (Element nextTurn : c_v ) {
				//	System.out.println(nextTurn.text());
				v.add(nextTurn.text());                           //list1
			}
			for (Element nextTurn : c_hp ) {
				//	System.out.println(nextTurn.text());
				hp.add(nextTurn.text());                           //list1
			}
			for (Element nextTurn : c_cc ) {
				//	System.out.println(nextTurn.text());
				cc.add(nextTurn.text());                           //list1
			}
			for (Element nextTurn : c_l ) {
				//	System.out.println(nextTurn.text());
				l.add(nextTurn.text());                           //list1
			}
			for (Element nextTurn : c_u ) {
				//	System.out.println(nextTurn.text());
				u.add(nextTurn.text());                           //list1
			}
			for (Element nextTurn : c_ss ) {
				//	System.out.println(nextTurn.text());
				ss.add(nextTurn.text());                           //list1
			}

			c_as.clear();//*
			c_lc.clear();
			c_nm.clear();//*
			c_v.clear();//*
			c_hp.clear();//*
			c_cc.clear(); //*
			c_l.clear(); //*
			//			 c_ei.clear();
			c_u.clear(); //*
			c_ss.clear();//*
			c_mcm.clear();



			int k=0,l1=0,m=0,n1=0,o=0,p=0,q=0,r=0,s=0;
			System.out.println("File size:"+as.size());
			System.out.println("temp:"+temp);
			for (int j = temp; j <=(as.size()+(temp-1)); j++) {



				row = sheet.createRow(j);
				for (int f = 0; f < head.size(); f++) { //0,1,2,3,4,5,6,7

					cell = row.createCell(f);
					switch (f) {
					case 0://list

						cell.setCellValue((String) as.get(k++));
						break;
					case 1://list1
						//cell = row.createCell(columnCount++);
						cell.setCellValue((String) nm.get(l1++));
						break;
					case 2://list2
						//cell = row.createCell(columnCount++);
						cell.setCellValue((String) v.get(m++));
						break;
					case 3://list2
						//cell = row.createCell(columnCount++);
						cell.setCellValue((String) hp.get(n1++));
						break;
					case 4://list2
						//cell = row.createCell(columnCount++);
						cell.setCellValue((String) cc.get(o++));

						//						cell.setCellStyle(stylewrap);
						break;
					case 5://list2
						//cell = row.createCell(columnCount++);
						cell.setCellValue((String) l.get(p++));
						break;
					case 6://list2
						//cell = row.createCell(columnCount++);
						cell.setCellValue((String) u.get(q++));
						break;
					case 7://list2
						//cell = row.createCell(columnCount++);
						cell.setCellValue((String) ss.get(r++));
						break;
					case 8://list2
						//cell = row.createCell(columnCount++);
						cell.setCellValue((String) filename);
						break;
					default:
						break;
					}

				}
			}
			temp=as.size()+temp;
			System.out.println("Total components:"+(temp-1));

			/*	workbook.write(outputStream);
			outputStream.close();
			workbook.close();*/
		}

		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);
		sheet.autoSizeColumn(3);
		sheet.autoSizeColumn(4);
		sheet.autoSizeColumn(5);
		sheet.autoSizeColumn(6);
		sheet.autoSizeColumn(7);
		sheet.autoSizeColumn(8);

		sheet.setAutoFilter(new CellRangeAddress(0,temp , 0, 8));
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();

		scan.close();
	}


	//Class-Path: jsoup-1.10.1.jar,poi-3.15.jar,poi-ooxml-3.15.jar,poi-ooxml-schemas-3.15.jar,xmlbeans-2.6.0.jar,commons-collections4-4.1.jar


}
/**
 * Author RAD9KOR_11_nov_2016
 * 
 * */
