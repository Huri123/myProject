package Test;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class ExcelRead {
    public static void main(String[] args) throws Exception{
File file=new File("src\\SampleData.xlsx");
        //System.out.println(file.exists());
        FileInputStream fileInputStream=new FileInputStream(file);

        ///workbook > sheet > roe > cell
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);

        XSSFSheet sheet=workbook.getSheet("Employees");

        System.out.println(sheet.getRow(2).getCell(1));

        int usedRows = sheet.getPhysicalNumberOfRows();

        int lastUsedRow = sheet.getLastRowNum();

        for (int rowNum=0;rowNum<usedRows;rowNum++){
            if (sheet.getRow(rowNum).getCell(0).toString().equals("Neena")){
                System.out.println("Neena's name: "+sheet.getRow(rowNum).getCell(0));
            }
        }

        for (int rowNum=0;rowNum<usedRows;rowNum++){
            if (sheet.getRow(rowNum).getCell(0).toString().equals("Adam")){
                System.out.println("Adam's lastname: "+sheet.getRow(rowNum).getCell(1));
            }
        }

        for (int rowNum=0;rowNum<usedRows;rowNum++){
            if (sheet.getRow(rowNum).getCell(1).toString().equals("King")){
                System.out.println("King's JOB_ID: "+sheet.getRow(rowNum).getCell(2));
            }
        }
    }
}
