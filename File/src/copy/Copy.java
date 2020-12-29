
package copy;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Copy {

	public static void main(String[] args) {

		  ArrayList<String> array = new ArrayList<String>();
		  int start = 0;
		  int count = 0;
		 
		  
		  
	
		try {
			Scanner sc = new Scanner(System.in);
			String excellPath = "";  
			  if(start == 0)
				{
				while(true)
				{
			    start = 0;
			   
			    System.out.print("엑셀 경로 입력 >>>>>");
			    String str = sc.next();
			   
				if(str.equals("q"))
				{
					System.exit(0);
				}
				if(str.equals("s"))
				{					
					start =1;
					break;
				}			
				array.add(str);
				 System.out.println("입력값: " + array.get(count));
				 count++;
				}
				} 
			
			  for(int i = 0 ; i < array.size(); i++)
			  {
				  System.out.println(array.size());
			 excellPath = "C:/Users/st3169/Desktop/확장자 테스트/" + array.get(i) + ".xlsx";
			FileInputStream file = new FileInputStream(excellPath);
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			int rowindex = 0;
			int columnindex = 0;
			String oldFolder = "";
			String newFolder = "";

			XSSFSheet sheet = workbook.getSheetAt(0);

			int rows = sheet.getPhysicalNumberOfRows();
			for (rowindex = 1; rowindex < rows; rowindex++) {

				XSSFRow row = sheet.getRow(rowindex);
				if (row != null) {
					int cells = row.getPhysicalNumberOfCells();
					for (columnindex = 0; columnindex <= cells; columnindex++) {
						XSSFCell cell = row.getCell(columnindex);
						String value = "";
						if (cell == null) {
							continue;
						} else {
							switch (cell.getCellType()) {
							case XSSFCell.CELL_TYPE_FORMULA:
								value = cell.getCellFormula();
								break;
							case XSSFCell.CELL_TYPE_NUMERIC:
								value = cell.getNumericCellValue() + "";
								break;
							case XSSFCell.CELL_TYPE_STRING:
								value = cell.getStringCellValue() + "";
								break;
							case XSSFCell.CELL_TYPE_BLANK:
								value = cell.getBooleanCellValue() + "";
								break;
							case XSSFCell.CELL_TYPE_ERROR:
								value = cell.getErrorCellValue() + "";
								break;
							}
						}
						switch (columnindex) {
						case 0: // 
							oldFolder = "C:/원본/" + value;
							//oldFolder = "D:/원본/" + value;
							break;
						case 1: // �깉寃쎈줈
							newFolder = "C:/결과물/테스트/" + value;
							//newFolder = "D:/결과물/1차/" + value;
							break;
						}
					}
					File folder1 = new File(oldFolder);
					File folder2 = new File(newFolder);
					System.out.println("이동 전"+folder1.toString());
					System.out.println("이동 후"+folder2.toString());
					Copy.copyDir(folder1, folder2);
				}
			}
			System.out.println("복사완료");

		}
		}catch (

		Exception e) {
			e.printStackTrace();
		}
		}
	

	public static void copyDir(File sourceF, File targetF) {

		if (!targetF.exists()) {
			targetF.mkdirs();
		}
		File[] fList = sourceF.listFiles();

		for (File result : fList) {
			File oldDir = new File(result.getAbsolutePath());
			File newDir = new File(targetF.getAbsolutePath() + "\\" + result.getName());
			
			
			String fileNameSub = result.getName();
			
			if (fileNameSub.length() >= 3) {

				String check0 = fileNameSub; // 확장자명이 3개 String check2 =
				String check1 = fileNameSub; // 확장자명이 3개 String check2 =
				String check2 = fileNameSub;
				String check3 = fileNameSub; // 확장자명이 5개
				check0 = fileNameSub.substring(fileNameSub.length() - 3, fileNameSub.length());
				check1 = fileNameSub.substring(fileNameSub.length() - 4, fileNameSub.length());
				check2 = fileNameSub.substring(fileNameSub.length() - 5, fileNameSub.length());
				check3 = fileNameSub.substring(fileNameSub.length() - 6, fileNameSub.length());
				if (check0.contains(".")) {
					fileNameSub = fileNameSub.substring(fileNameSub.length() - 3, fileNameSub.length());
				} else if (check1.contains(".")) {

					fileNameSub = fileNameSub.substring(fileNameSub.length() - 4, fileNameSub.length());

				} else if (check2.contains(".")) {

					fileNameSub = fileNameSub.substring(fileNameSub.length() - 5, fileNameSub.length());

				} else if (check3.contains(".")) {

					fileNameSub = fileNameSub.substring(fileNameSub.length() - 6, fileNameSub.length());
				}
				System.out.println("확장자 명 "+ fileNameSub);
			}
			if (result.isDirectory()) {
				copyDir(oldDir, newDir);
			} else if (result.isFile()) {
				if(fileNameSub.equalsIgnoreCase(".tif") || fileNameSub.equalsIgnoreCase(".tiff") )
				{
					
				   copyFile(result, newDir);
				}
			}
		}
	}

	public static boolean copyFile(File oldFile, File newDir) {
		boolean result = false;
		if (oldFile.exists()) {
			try {
				FileInputStream fis = new FileInputStream(oldFile.getAbsolutePath());
				FileOutputStream fos = new FileOutputStream(newDir.getAbsolutePath());




				fis = new FileInputStream(oldFile);
				fos = new FileOutputStream(newDir);
				byte[] b = new byte[4096];
				int cnt = 0;
				while ((cnt = fis.read(b)) != -1) {
					fos.write(b, 0, cnt);
				}

				fis.close(); // �궗�슜�셿猷�
				fos.close(); // �궗�슜�셿猷�

				System.out.println(oldFile.getAbsolutePath() + "파일");
				System.out.println(newDir.toString()+ "위치로 복사 완료");
				
			} catch (Exception e) {
				result = false;
			}

		} else {
			System.out.println("복사완료");

		}
		return result;

	}

}
