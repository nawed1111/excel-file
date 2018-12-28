
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class excel {
	static int flag = 0;
    public static void main(String args[]) throws Exception{
    	FileInputStream f1 = new FileInputStream(new File("D:\\excel\\baseline_environment_1_Dev_abkx_20181205193434.xlsx"));
    	FileInputStream f2 = new FileInputStream(new File("D:\\excel\\baseline_environment_1_Dev_abkx_20181205193435.xlsx"));
    	
    	XSSFWorkbook wb1 = new XSSFWorkbook(f1);
    	XSSFWorkbook wb2 = new XSSFWorkbook(f2);
    	XSSFWorkbook wb3 =  wb2;

        compare(wb1, wb2, wb3);
        wb1.close();
        wb2.close();
    }
    
    public static void compare(XSSFWorkbook wb1, XSSFWorkbook wb2, XSSFWorkbook wb3) throws Exception{
        Locator loc1 = new Locator();
        Locator loc2 = new Locator();
        Locator loc3 = new Locator();
        loc1.workbook = wb1;
        loc2.workbook = wb2;
        loc3.workbook = wb3;
        
        compareDataInAllSheets(loc1, loc2, loc3);
        
        FileOutputStream out = new FileOutputStream(new File("D:\\excel\\Writesheet.xlsx"));
		wb3.write(out);
	    out.close();
    }

    private static void compareDataInAllSheets(Locator loc1, Locator loc2, Locator loc3){

        for (int i = 2; i < loc1.workbook.getNumberOfSheets(); i++){
            if (loc2.workbook.getNumberOfSheets() <= i) return;

            loc1.sheet = loc1.workbook.getSheetAt(i);
            loc2.sheet = loc2.workbook.getSheetAt(i);
            loc3.sheet = loc3.workbook.getSheetAt(i);

            compareDataInSheet(loc1, loc2, loc3);
        }
    }

    private static void compareDataInSheet(Locator loc1, Locator loc2, Locator loc3){	
    	
    	XSSFCellStyle style = loc3.workbook.createCellStyle();
    	XSSFFont font= loc3.workbook.createFont();
    	font.setFontHeightInPoints((short)9);
    	font.setFontName("Verdana");
    	font.setBold(true);
    	style.setFont(font); 
        for (int j = 0; j < loc1.sheet.getPhysicalNumberOfRows(); j++){	
            if (loc2.sheet.getPhysicalNumberOfRows() <= j) return;
           
            loc1.row = loc1.sheet.getRow(j);
            loc2.row = loc2.sheet.getRow(j);
            loc3.row = loc3.sheet.getRow(j);
            if ((loc1.row == null) || (loc2.row == null)) {
                continue;
               }
            int Count = loc3.row.getLastCellNum();
            if(j==0){
            	loc3.cell = loc3.row.getCell(5);
            	loc3.cell.setCellValue("Current CheckSum");
            	
            	loc3.cell = loc3.row.createCell(Count);
            	loc3.cell.setCellValue("Prev. CheckSum");
            	loc3.cell.setCellStyle(style);
            	loc3.sheet.autoSizeColumn(Count++);
            	
            	loc3.cell = loc3.row.createCell(Count);
            	loc3.cell.setCellValue("Status");
            	loc3.cell.setCellStyle(style);
            	loc3.sheet.autoSizeColumn(Count++);
            }
            else{
            	 compareDataInRow(loc1, loc2);
            	
            	loc1.cell = loc1.row.getCell(5);
            	String str = loc1.cell.getStringCellValue();
            	loc3.cell = loc3.row.createCell(Count++);     
            	loc3.cell.setCellValue(str);
            	
            	loc3.cell = loc3.row.createCell(Count++);
            	
            	if(flag == 0)
            		loc3.cell.setCellValue("Ok");
            	else{
            		loc3.cell.setCellValue("Not Ok");
            		flag =0;
            	}
            }
         }
    }
    private static void compareDataInRow(Locator loc1, Locator loc2){	 
    	
    	for (int k = 0; k < loc1.row.getPhysicalNumberOfCells(); k++){ 
    	 if (loc2.row.getPhysicalNumberOfCells()<= k) break;
            loc1.cell = loc1.row.getCell(k);
            loc2.cell = loc2.row.getCell(k);

            if ((loc1.cell == null) || (loc2.cell == null)){
            	continue;
            }
           compareDataInCell(loc1, loc2);
    	}
    }

    private static void compareDataInCell(Locator loc1, Locator loc2){
        if (isCellTypeMatches(loc1, loc2)){
            final CellType loc1cellType = loc1.cell.getCellType();
            switch(loc1cellType){
                case BLANK:
                case STRING:
                case ERROR:
                    	isCellContentMatches(loc1,loc2);
                    	break;
                case NUMERIC:
                	    isCellContentMatchesForNumeric(loc1,loc2);
                        break;
                default:
                    	throw new IllegalStateException("Unexpected cell type: " + loc1cellType);
            }
        }
    }
    private static boolean isCellTypeMatches(Locator loc1, Locator loc2){
        CellType type1 = loc1.cell.getCellType();
        CellType type2 = loc2.cell.getCellType();
        if (type1 == type2) return true;
        return false;
    } 
    private static void isCellContentMatches(Locator loc1, Locator loc2){
        String str1 = loc1.cell.getRichStringCellValue().getString();
        String str2 = loc2.cell.getRichStringCellValue().getString();
        if (!str1.equals(str2)) 
           flag++;
    }

    private static void isCellContentMatchesForNumeric(Locator loc1, Locator loc2){
        double num1 = loc1.cell.getNumericCellValue();
        double num2 = loc2.cell.getNumericCellValue();
        if (num1 != num2) 
        	flag++;
        }
    private static class Locator{
        XSSFWorkbook workbook;
        XSSFSheet sheet;
        XSSFRow row;
        Cell cell;
    }
}

