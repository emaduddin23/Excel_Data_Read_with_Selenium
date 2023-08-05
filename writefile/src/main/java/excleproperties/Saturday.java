package excleproperties;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class Saturday {
@Test
	public void test() throws IOException{
	    String excelFilePath = "C:\\Users\\EMAD\\eclipse-workspace\\writefile\\src\\main\\java\\datafiles\\Excel.xlsx";
	    FileInputStream inputStream = new FileInputStream(excelFilePath);

	    XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	    XSSFSheet sheet = workbook.getSheetAt(0);
	    
	    int rows=sheet.getLastRowNum();
	    int cols =sheet.getRow(1).getLastCellNum();

	    for(int r=2 ; r<rows;r++)
	    {
	        XSSFRow row = sheet.getRow(r);
	        for(int c=1; c<cols;c++)
	        {
	            XSSFCell cell=row.getCell(c);
	            switch(cell.getCellType())
	            {
	                case STRING: System.out.print(cell.getStringCellValue()); break;
	                case NUMERIC: System.out.print(cell.getNumericCellValue());break;
	                case BOOLEAN: System.out.print(cell.getBooleanCellValue());break;
				default:
					break;

	            }
	            System.out.print(" | ");
	        }
	        System.out.println();
	    }
	
	}



}
