import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Date;
import java.sql.RowId;
import java.util.Set;

import org.w3c.dom.stylesheets.StyleSheet;



public class fileoperation {
	
	public static void main (String []args)throws IOException{
		File newfile = new File("C:\\Users\\home\\eclipse-workspace\\SeleniumTesting\\target\\Grid_Tie_Without_Battery_Design_Simulation_Tool (1).xlsx");
		
		FileInputStream t = new FileInputStream(newfile);
		
		HSSFWorkbook work = new HSSFWorkbook(t);
		
		Sheet a = work.getSheet("sheet1");
		
		for(int i = 0; i<a.getPysicalNumberOfRows(); i++) {
			
			Row row = a.getRow(i);
			
			for (int j = 0; i < row.getPhysicalNumberOfCells(); j++) 
			{
				
			Cell cell = row.getCell(j);
			int cellType = cell.getCellType();
			if(cellType==1) {
				String value= cell.getStringCellValue();
				System.out.println("Name: "+value);
				
				
				
				
			}else {
				if{ Date.iscellDateFormotted(cell)){
					Date d = cell.getCellValue();
					SimpleDateFormate date = new SimpleDateFormate("dd/MM/yy");
					System.out.println("DOB: "+value);
					System.out.println("");
				}
				else{
					long value = (long)cell.getNumericCellValue();
					System.out.println("Age: "+value);
				}
				}
					
					
					
				}
				
				}
			}
			
			

			
			}
		}
		
		
		
		
		
		
	}

}
