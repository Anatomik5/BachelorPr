package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Main {

    public static void main(String[] args) throws Exception{
        File myfile = new File("C:/savedexcel/test.xlsx");
        FileInputStream file = new FileInputStream(myfile);
        XSSFWorkbook workbook = new XSSFWorkbook(file);

        XSSFSheet spreadsheet = workbook.getSheet("Student Data"); //just name of table

        XSSFRow row; // creating a row

        Map<Integer, Object[]> studentData = new TreeMap<Integer, Object[]>();

        studentData.put(1, new Object[] { "11111111", "Simon", "10"});

        studentData.put(2, new Object[] { "21e4124", "Johnny", "11"});

        Set<Integer> keyid = studentData.keySet();

        int rowid = spreadsheet.getPhysicalNumberOfRows();

        for (Integer key : keyid) { // writing the data into the spreadsheet
            row = spreadsheet.createRow(rowid++); //creating new row in table
            Object[] objectArr = studentData.get(key);
            int cellid = 0;
            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }
        FileOutputStream out = new FileOutputStream(new File("C:/savedexcel/test.xlsx"));
        workbook.write(out);
        file.close();
        out.close();
    }
}