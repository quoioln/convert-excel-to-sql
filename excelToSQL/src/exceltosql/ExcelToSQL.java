package exceltosql;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author grega vrabancic
 */
public class ExcelToSQL {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        String filename = "city_post.xlsx";
        
    }
    
    public static void convertToSQL(String filename) throws Exception{
        String[] attributes=new String[2];
        CityPost[] cityPost=readSheet(filename, attributes);
    }
    
    public static CityPost[] readSheet(String filename, String[] attributes) throws Exception{
        FileInputStream fis = null;
        CityPost[] cityPosts=null;
        try {
            fis = new FileInputStream(filename);
            
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
            int number=sheet.getLastRowNum();
            System.out.println(" number of rows: "+ number);
            System.out.println("Loading...");
            boolean firstRow=true;
            Integer counter=0;
            Integer counterRows=-1;
            cityPosts=new CityPost[number];
            while (rows.hasNext())
            {
                XSSFRow row = ((XSSFRow) rows.next());
                Iterator cells = row.cellIterator();
                while(cells.hasNext())
                {
                    XSSFCell cell = (XSSFCell) cells.next();
                    if(firstRow){
                        attributes[counter]=cell.getStringCellValue();
                    }
                    else{
                        cityPosts[counterRows].setCity(cell.getStringCellValue());
                        cityPosts[counterRows].setPost(Integer.parseInt(cell.getRawValue()));
                    }
                }
                firstRow=false;
                counter++;
                counterRows++;
             }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                fis.close();
            }
        }
        return cityPosts;
    }
}