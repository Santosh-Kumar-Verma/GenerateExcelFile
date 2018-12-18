package com.test.action;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.test.dto.Employee;

public class ExcelOerationUtil {

	public static void writeFileToXml(String desn,List<Employee> empList) {
		 
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Employee Deatils");
        
        Row headingRow = sheet.createRow(0);
        int headinCell = 0;
        Cell cell = headingRow.createCell(headinCell++);
        cell.setCellValue("EmployeeId");
        Cell empNameCell = headingRow.createCell(headinCell++);
        empNameCell.setCellValue("EmployeeName");
        Cell mobleNoCell = headingRow.createCell(headinCell++);
        mobleNoCell.setCellValue("MobileNo");
        
        int rowNum = 1;
        for (Employee employee : empList) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            Cell cell1 = row.createCell(colNum++);
            cell1.setCellValue(employee.getEmployeeId());
            Cell cell2 = row.createCell(colNum++);
            cell2.setCellValue(employee.getEmpName());
            Cell cell3 = row.createCell(colNum++);
            cell3.setCellValue(employee.getMobileNo());
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(desn);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    
	}
	public static void readFileToXml(String desn) {
		try {

            FileInputStream excelFile = new FileInputStream(new File(desn));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    //getCellTypeEnum shown as deprecated for version 3.15
                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.print(currentCell.getStringCellValue() + "  ");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        System.out.print(currentCell.getNumericCellValue() + "  ");
                    }

                }
                System.out.println();

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
	}
}
