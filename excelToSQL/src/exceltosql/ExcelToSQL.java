package exceltosql;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.io.FileWriter;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author grega vrbancic
 */
public class ExcelToSQL {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        String filename = "city_post.xlsx";
        String db;
        Scanner scan = new Scanner(System.in);
        System.out.print("Database name: ");
        db=scan.nextLine();
        
        try {
            convertToSQL(filename, db);
        } catch (Exception ex) {
            Logger.getLogger(ExcelToSQL.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public static void convertToSQL(String filename, String db) throws Exception{
        String[] attributes=new String[2];
        if(readSheet(filename, attributes, db)){
            System.out.println("Convert done!");
        }
    }
    
    public static boolean readSheet(String filename, String[] attributes, String db) throws Exception{
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(filename);
            FileWriter fstream = new FileWriter("SQL.sql");
            BufferedWriter out = new BufferedWriter(fstream);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
            int number=sheet.getLastRowNum();
            System.out.println(" number of rows: "+ number);
            System.out.println("Loading...");
            boolean firstRow=true;
            Integer counter=0;
            String city;
            Integer post=0;
            while (rows.hasNext())
            {
                XSSFRow row = ((XSSFRow) rows.next());
                Iterator cells = row.cellIterator();
                while(cells.hasNext())
                {
                    XSSFCell cell = (XSSFCell) cells.next();
                    if(firstRow){
                        attributes[counter]=cell.getStringCellValue();
                        counter++;
                    }
                    else{
                        city=cell.getStringCellValue();
                        if(cells.hasNext()){
                           cell = (XSSFCell) cells.next();
                           post = (int)cell.getNumericCellValue();   
                        }
                        out.write("INSERT INTO " + db + ".city_post (" + attributes[0] + ", "
                                + attributes[1] + ") VALUES ('" + city + "', " + post +");\n");
                    }
                }
                firstRow=false;               
             }
            fstream.close();
            return true;
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                fis.close();
            }
        }
        return false;
    }
}