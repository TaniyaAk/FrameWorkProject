package com.taniya.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLOutput;

//how to read data from EXCEL?
public class ExcelReader {
    public static void main(String[] args) throws IOException {
        printExcelData(readExcelData());

    }

    public static String[][] readExcelData() throws IOException {
        String filePath = "C:\\Users\\19176\\Desktop\\testdata.xlsx";
        FileInputStream stream = new FileInputStream(filePath);

        //Loading excel file
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(stream);//(one whole file)This is from apache.poi.
        //Isolating datasheet
        Sheet dataSheet = xssfWorkbook.getSheetAt(0);//One data sheet

        //Line 25-27: Created to DRA, that is dynamic based on the data sheets,row counts and column counts
        int rowCount = dataSheet.getLastRowNum();
        int columnCount = dataSheet.getRow(0).getLastCellNum();
        String[][] data = new String[rowCount][columnCount];

        for (int i = 1; i <= rowCount; i++) { /*we started from index ONE, not from header(0)*/
            //Isolating data row
            Row row = dataSheet.getRow(i); //after to DRA, we are accessing the Row of each Row

            for (int j = 0; j < rowCount; j++) { /*after that we are creating forLoop, & in the forLoop we are going to
                the each of the data rows.  */
                //Isolating the cell
                Cell cell = row.getCell(j);
                /*purpose of try & catch is like handing an Exception.
                if theres no data in the cell we wont be able to get any string value.
                if theres no data or string value, we put something in the "TODRA",
                if the data is not string type, the "TODRA(dynamic based on the datasheet)" will not except that and it will
                throw exception to stop executing your code. To prevent that we
                have created the "try catch" block. It will still throw the Exception.
                But your code will continue to run the next cell, cause "try catch" will handle that.*/
                try {
                    //Entering data
                    data[i - 1][j] = cell.getStringCellValue();
                } catch (Exception exception) {
                    exception.printStackTrace();
                }

            }
        }

        return data;
    }

    public static void printExcelData(String[][] theData) {
        for (int i = 0; i < theData.length; i++) {
            for (int j = 0; j < theData[i].length; j++) {
                System.out.println(theData[i][i]);
            }
        }
    }
}
