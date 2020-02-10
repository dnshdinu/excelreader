package read;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readexl {
	
	public static void main(String[] args) {

	    try {
	    	File loc=new File("C:\\Users\\Dineshkumar\\eclipse-workspace\\Readexcel\\excel\\Data.xlsx");
			FileInputStream stream = new FileInputStream(loc);
		  try {
			Workbook w=new XSSFWorkbook(stream);
			Sheet s=w.getSheet("Sheet1");
		      for (int i =0; i <=s.getPhysicalNumberOfRows(); i++) {
			  Row r = s.getRow(i);
			  for (int j = 0; j <=r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
			
					int type = c.getCellType();
					

					if(type==1) {
					
					  String	name=c.getStringCellValue();
						System.out.println(name);
					}
					else if (type==0) {
						if (DateUtil.isCellDateFormatted(c)) {
						String	name=new SimpleDateFormat("dd-mm-yy").format(c.getDateCellValue());
							System.out.println(name);
						} else {

						String	name=String.valueOf(c.getNumericCellValue());
							System.out.println(name);
						}
						

					}
				}
				
			}

			
		} catch (IOException e) {

			e.printStackTrace();
		}
			
			
			
			
		} catch (FileNotFoundException e) {

			e.printStackTrace();
		}
	}
	
	
	

}
