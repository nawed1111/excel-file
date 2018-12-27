import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellReference;



public class ExcelComparator {
    
    private static final String CELL_DATA_DOES_NOT_MATCH = "Cell Data does not Match ::";
    static List<String> listOfDifferences = new ArrayList<String>();
  
    private static class Locator 
    {
        XSSFWorkbook workbook;
        XSSFSheet sheet;
        XSSFRow row;
        Cell cell;
    }
    public static void main(String args[]) throws Exception 
    {
    	FileInputStream f1 = new FileInputStream(new File("D:\\excel\\baseline_environment_1_Dev_abkx_20181205193434.xlsx"));
    	FileInputStream f2 = new FileInputStream(new File("D:\\excel\\baseline_environment_1_Dev_abkx_20181205193435.xlsx"));
        
    	XSSFWorkbook wb1 = new XSSFWorkbook(f1);
    	XSSFWorkbook wb2 = new XSSFWorkbook(f2);
        compare(wb1, wb2);
        
        wb1.close();
        wb2.close();
        excelCreate();
    }
    
    public static List<String> compare(XSSFWorkbook wb1, XSSFWorkbook wb2) 
    {
        Locator loc1 = new Locator();
        Locator loc2 = new Locator();
        loc1.workbook = wb1;
        loc2.workbook = wb2;
        ExcelComparator excelComparator = new ExcelComparator();
        excelComparator.compareNumberOfSheets(loc1, loc2 );
        excelComparator.compareSheetNames(loc1, loc2);
        excelComparator.compareSheetData(loc1, loc2);
        return excelComparator.listOfDifferences;
    }

    private void compareDataInAllSheets(Locator loc1, Locator loc2) 
    {

        for (int i = 2; i < loc1.workbook.getNumberOfSheets(); i++) 
        {
            if (loc2.workbook.getNumberOfSheets() <= i) return;

            loc1.sheet = loc1.workbook.getSheetAt(i);
            loc2.sheet = loc2.workbook.getSheetAt(i);

            compareDataInSheet(loc1, loc2);
        }
    }

    private void compareDataInSheet(Locator loc1, Locator loc2) 
    {	int k[] =cell(loc1);
        for (int j = 0; j < loc1.sheet.getPhysicalNumberOfRows(); j++) 
        {
            if (loc2.sheet.getPhysicalNumberOfRows() <= j) return;
            loc1.row = loc1.sheet.getRow(j);
            loc2.row = loc2.sheet.getRow(j);
            
            if ((loc1.row == null) || (loc2.row == null)) {
             continue;
            }
            for (int l = 0; l < k.length; l++) 
            {
            	compareDataInRow(loc1, loc2, k[l]);
           }
        }
    }
    private void compareDataInRow(Locator loc1, Locator loc2, int l) 
    {	 	
            loc1.cell = loc1.row.getCell(l);
            loc2.cell = loc2.row.getCell(l);

            if ((loc1.cell == null) || (loc2.cell == null)) 
            {
            	return;
            	
            }
            compareDataInCell(loc1, loc2);
    }

    private void compareDataInCell(Locator loc1, Locator loc2) 
    {
        if (isCellTypeMatches(loc1, loc2)) 
        {
            final CellType loc1cellType = loc1.cell.getCellType();
            switch(loc1cellType) 
            {
                case BLANK:
                case STRING:
                case ERROR:
                    isCellContentMatches(loc1,loc2);
                    break;
                case BOOLEAN:
                    isCellContentMatchesForBoolean(loc1,loc2);
                    break;
                case FORMULA:
                    isCellContentMatchesForFormula(loc1,loc2);
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(loc1.cell)) 
                    {
                        isCellContentMatchesForDate(loc1,loc2);
                    } 
                    else 
                    {
                        isCellContentMatchesForNumeric(loc1,loc2);
                    }
                    break;
                default:
                    throw new IllegalStateException("Unexpected cell type: " + loc1cellType);
            }
        }
    }
    private void compareNumberOfColumnsInSheets(Locator loc1, Locator loc2) {
        for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
            if (loc2.workbook.getNumberOfSheets() <= i) return;
            
            loc1.sheet = loc1.workbook.getSheetAt(i);
            loc2.sheet = loc2.workbook.getSheetAt(i);

            Iterator<Row> ri1 = loc1.sheet.rowIterator();
            Iterator<Row> ri2 = loc2.sheet.rowIterator();
            
            int num1 = (ri1.hasNext()) ? ri1.next().getPhysicalNumberOfCells() : 0;
            int num2 = (ri2.hasNext()) ? ri2.next().getPhysicalNumberOfCells() : 0;
            
            if (num1 != num2) {
                String str = String.format(Locale.ROOT, "%s\nworkbook1 -> %s [%d] != workbook2 -> %s [%d]",
                    "Number Of Columns does not Match ::",
                    loc1.sheet.getSheetName(), num1,
                    loc2.sheet.getSheetName(), num2
                );
                listOfDifferences.add(str);
            }
        }
    }

    private void compareNumberOfRowsInSheets(Locator loc1, Locator loc2) {
        for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
            if (loc2.workbook.getNumberOfSheets() <= i) return;

            loc1.sheet = loc1.workbook.getSheetAt(i);
            loc2.sheet = loc2.workbook.getSheetAt(i);
            
            int num1 = loc1.sheet.getPhysicalNumberOfRows();
            int num2 = loc2.sheet.getPhysicalNumberOfRows();

            if (num1 != num2) {
                String str = String.format(Locale.ROOT, "%s\nworkbook1 -> %s [%d] != workbook2 -> %s [%d]",
                    "Number Of Rows does not Match ::",
                    loc1.sheet.getSheetName(), num1,
                    loc2.sheet.getSheetName(), num2
                );
                listOfDifferences.add(str);
            }
        }

    }
    private void compareNumberOfSheets(Locator loc1, Locator loc2) {
        int num1 = loc1.workbook.getNumberOfSheets();
        int num2 = loc2.workbook.getNumberOfSheets();
        if (num1 != num2) {
            String str = String.format(Locale.ROOT, "%s\nworkbook1 [%d] != workbook2 [%d]",
                "Number of Sheets do not match ::",
                num1, num2
            );

            listOfDifferences.add(str);
            
        }
    }
    private void compareSheetData(Locator loc1, Locator loc2) {
        compareNumberOfRowsInSheets(loc1, loc2);
        compareNumberOfColumnsInSheets(loc1, loc2);
        compareDataInAllSheets(loc1, loc2);

    }
    private void compareSheetNames(Locator loc1, Locator loc2) {
        for (int i = 0; i < loc1.workbook.getNumberOfSheets(); i++) {
            String name1 = loc1.workbook.getSheetName(i);
            String name2 = (loc2.workbook.getNumberOfSheets() > i) ? loc2.workbook.getSheetName(i) : "";
            
            if (!name1.equals(name2)) {
                String str = String.format(Locale.ROOT, "%s\nworkbook1 -> %s [%d] != workbook2 -> %s [%d]",
                    "Name of the sheets do not match ::", name1, i+1, name2, i+1
                );
                listOfDifferences.add(str);
            }
        }
    }
    private boolean isCellTypeMatches(Locator loc1, Locator loc2) 
    {
        CellType type1 = loc1.cell.getCellType();
        CellType type2 = loc2.cell.getCellType();
        if (type1 == type2) return true;
        addMessage(loc1, loc2, "Cell Data-Type does not Match in :: ", type1.name(), type2.name());
        return false;
    }

    private void addMessage(Locator loc1, Locator loc2, String messageStart, String value1, String value2) 
    {
        String str =
            String.format(Locale.ROOT, "%s\nworkbook1 -> %s -> %s [%s] != workbook2 -> %s -> %s [%s]", messageStart, loc1.sheet.getSheetName(), new CellReference(loc1.cell).formatAsString(), value1, loc2.sheet.getSheetName(), new CellReference(loc2.cell).formatAsString(), value2);
        listOfDifferences.add(str);
    }
    
    private void isCellContentMatches(Locator loc1, Locator loc2) 
    {
        String str1 = loc1.cell.getRichStringCellValue().getString();
        String str2 = loc2.cell.getRichStringCellValue().getString();
        if (!str1.equals(str2)) 
        {
            addMessage(loc1,loc2,CELL_DATA_DOES_NOT_MATCH,str1,str2);
        
        }
    }
    
    
    private void isCellContentMatchesForBoolean(Locator loc1, Locator loc2) 
    {
        boolean b1 = loc1.cell.getBooleanCellValue();
        boolean b2 = loc2.cell.getBooleanCellValue();
        if (b1 != b2) 
        {
            addMessage(loc1,loc2,CELL_DATA_DOES_NOT_MATCH,Boolean.toString(b1),Boolean.toString(b2));
        }
    }
    
    
    private void isCellContentMatchesForDate(Locator loc1, Locator loc2) 
    {
        Date date1 = loc1.cell.getDateCellValue();
        Date date2 = loc2.cell.getDateCellValue();
        if (!date1.equals(date2)) 
        {
            addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, date1.toString(), date2.toString());
           
        }
    }
    
    private void isCellContentMatchesForFormula(Locator loc1, Locator loc2) 
    {
        String form1 = loc1.cell.getCellFormula();
        String form2 = loc2.cell.getCellFormula();
        if (!form1.equals(form2)) 
        {
            addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, form1, form2);
           
        }
    }

    private void isCellContentMatchesForNumeric(Locator loc1, Locator loc2) 
    {
        double num1 = loc1.cell.getNumericCellValue();
        double num2 = loc2.cell.getNumericCellValue();
        if (num1 != num2) 
            addMessage(loc1, loc2, CELL_DATA_DOES_NOT_MATCH, Double.toString(num1), Double.toString(num2));
    }
   private static void checkList()
   {
   	if(listOfDifferences.size()==0)
   	{
   		listOfDifferences.add("No Difference in two files.");
    	}
    }
   private int[] cell(Locator loc1) {
		 
		 Scanner sc = new Scanner(System.in);
		
		 System.out.println("Enter the number of cells to compare:");
		 int n = sc.nextInt();
		 int k[] = new int[n];
		 for(int i = 0 ; i<n ; i++)
		 {
	    System.out.println("Enter the  Column to compare:");
	    String s = sc.next();
	    CellReference cellReference = new CellReference(s);
	    Cell cell = loc1.sheet.getRow(0).getCell(cellReference.convertColStringToIndex(s));
	    k[i] = cell.getColumnIndex();
	    System.out.println(k[i]);
		 }
	    sc.close();
	    return k;
	   }
    private static void excelCreate()
    {
    	XSSFWorkbook workbook = new XSSFWorkbook(); 
    	XSSFSheet spreadsheet = workbook.createSheet("Comparison_info");
    	spreadsheet.protectSheet("password");
    	XSSFRow row;
    	checkList();
    	for(int i = 0 ; i<listOfDifferences.size();i++)
    	{	
    		row = spreadsheet.createRow(i);
    		Cell cell = row.createCell(0);
    		cell.setCellValue(listOfDifferences.get(i));
    		spreadsheet.autoSizeColumn(i);
    	}
    	FileOutputStream out;
    	try 
    	{
    		out = new FileOutputStream(new File("D:\\excel\\Writesheet.xlsx"));
    		workbook.write(out);
		    out.close();
		    workbook.close();
		    System.out.println("Writesheet.xlsx written successfully");
    	} 
    	catch (Exception e) 
    	{
    		e.printStackTrace();
    	}
    }
}

