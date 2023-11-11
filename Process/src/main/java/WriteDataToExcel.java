import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WriteDataToExcel {

    public static void main(String[] args) throws Exception{
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet spreadsheet = workbook.createSheet("Student Data"); //just name of table

        XSSFRow row; // creating a row

        Map<Integer, Object[]> studentData = new TreeMap<Integer, Object[]>();


        studentData.put(1, new Object[] { "Inv.Nr.", "Alte Inv.Nr.", "Schublade"});

        studentData.put(2, new Object[] { "11111111", "Simon", "10"});

        studentData.put(3, new Object[] { "11111112", "Johnny", "11"});

        Set<Integer> keyid = studentData.keySet();

        int rowid = 0;

        for (Integer key : keyid) { // writing the data into the spreadsheet
            row = spreadsheet.createRow(rowid++); //creating new row in table
            Object[] objectArr = studentData.get(key);
            int cellid = 0;
            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }

        FileOutputStream out = new FileOutputStream("C:/savedexcel/test.xlsx"); // writing in existing excel, change it for new directory
        workbook.write(out); //push file
        out.close();
    }
}
