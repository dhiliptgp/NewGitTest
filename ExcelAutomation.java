package TestPackage;

import org.testng.annotations.Test;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

@SuppressWarnings("unused")
public class ExcelAutomation {

	@Test
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		
    FileInputStream excel=new FileInputStream("C:\\Users\\user1\\Desktop\\MyFamily.xlsx");

    XSSFWorkbook Wb = new XSSFWorkbook(excel);
    
    Sheet mysheet=Wb.getSheet("Sheet1");
    int rowcount=mysheet.getLastRowNum();
    
    
    
    for (int i=0;i<rowcount;i++ ) {
    	
    	Row row=mysheet.getRow(i);
    	System.out.println(i);
    	for(int j=0;j<=row.getLastCellNum();j++) {
    		String text=row.getCell(j).getStringCellValue();
    		System.out.println("Row "+i+" Column "+j+" value is "+ text);
    	}
    	
    	
    }
    Wb.close();
    excel.close();
    
	}

}
