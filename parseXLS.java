import java.io.File;

import jxl.Sheet;
import jxl.Workbook;

public class parseXLS{
	
	/* getDataFromXLS method returns 2D array of object data type which represent rows and columns in the XLS file.
	 * getDataFromXLS method has 4 arguments:
	 * 1) String: file -> XLS file directory
	 * 2) String: sheetName -> name of the sheet in the XLS file
	 * 3) cols: int -> represents the number of columns you would like to be parsed
	 * 4) startRow: int -> from which row you would like to parse the sheet
	 */
	public static Object[][] getDataFromXLS(String file, String sheetName, int cols, int startRow) throws Exception{
		startRow --;
		//open the XLS file
		Workbook workbook = Workbook.getWorkbook(new File(file));
		//point to specific sheet
		Sheet sheet = workbook.getSheet(sheetName);
		//get number of rows minus starting row
		int records = sheet.getRows() - startRow;
		int currentPosition = startRow;
		//create 2D array of object data type
		Object[][] values = new Object[records][cols];
		//loop over the rows and columns to parse value by value
		for(int i = 0 ; i < records ; i++, currentPosition++){
			for(int j = 0 ; j < cols ; j++) values[i][j] = sheet.getCell(j, currentPosition).getContents();
		}
		//close the XLS file
		workbook.close();
		//return the parsed values
		return values;
	}

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Object[][] values = getDataFromXLS("xls_file.xls", "excel_sheet", 3, 2);
		for(int i=0 ; i < values.length ; i++){
			for(int j=0 ; j < values[i].length ; j++){
				System.out.println(values[i][j]);
			}
			System.out.println("====================================");
		}
	}

}
