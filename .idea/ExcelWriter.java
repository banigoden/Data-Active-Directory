import org.apache.commons.collections4.iterators.EmptyListIterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.naming.NamingException;
import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;

public class ExcelWriter {



    public void ExcelWrite(String sheetName, ArrayList<Employee> employees) throws IOException {
        XSSFWorkbook workbook; // new HSSFWorkbook() for generating `.xls` file
        XSSFSheet sheet;
        File file = new File("C:\\Users\\dbandurin\\IdeaProjects\\contactAD\\src\\main\\ContactAD.xlsx");

        if(!file.exists()){
            workbook = new XSSFWorkbook();
            sheet =workbook.createSheet(sheetName);
        }else{
            FileInputStream read = new FileInputStream(file);
            workbook = new XSSFWorkbook(read);
            sheet =workbook.createSheet(sheetName);

        }

        try {
            //Get the count in sheet
            for (int i = 0; i < 6 ; i++) {
                sheet.setColumnWidth(i, 10000);
            }

            int rowCount = sheet.getLastRowNum();
            Row empRow = sheet.createRow(0);
            Cell c0 = empRow.createCell(0);
            c0.setCellValue("Name");
            Cell c5 = empRow.createCell(5);
            c5.setCellValue("Department");
            Cell c4 = empRow.createCell(4);
            c4.setCellValue("Title");
            Cell c3 = empRow.createCell(3);
            c3.setCellValue(("Mail"));
            Cell c1 = empRow.createCell(1);
            c1.setCellValue("Telephone");
            Cell c2 = empRow.createCell(2);
            c2.setCellValue("Mobile");

        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        sheet = workbook.getSheetAt(workbook.getSheetIndex(sheetName));
        for (int count = 1; count < employees.size() ; count++) {

            Row empRow = sheet.createRow(count);

            Cell c0 = empRow.createCell(0);
            c0.setCellValue(employees.get(count).getName());
            Cell c5 = empRow.createCell(5);
            c5.setCellValue(employees.get(count).getDepartment());
            Cell c4 = empRow.createCell(4);
            c4.setCellValue(employees.get(count).getTitle());
            Cell c3 = empRow.createCell(3);
            c3.setCellValue(employees.get(count).getEmail());

            Cell c1 = empRow.createCell(1);
            c1.setCellValue(employees.get(count).getNumber());
            Cell c2 = empRow.createCell(2);
            c2.setCellValue(employees.get(count).getMobile());

        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new
                    File("C:\\Users\\dbandurin\\IdeaProjects\\contactAD\\src\\main\\ContactAD.xlsx"));
            workbook.write(out);
            out.close();
            //System.out.println("Update Successfully");
        }catch (Exception e) {
            System.out.println("Did't write " + e);
        }
    }


}
