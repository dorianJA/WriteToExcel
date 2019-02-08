package Test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class TestExcel {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook(new FileInputStream("D:\\Desktop.xlsx"));
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\myNew.xlsx");
        Workbook workbookNew = new XSSFWorkbook();
        Sheet sheet = workbookNew.createSheet();

        int countRow = 0;
        int countCell;
        for(Row row: workbook.getSheetAt(0)){
               Row row1 = sheet.createRow(countRow);
                countRow++;
                countCell = 0;

            for(Cell cell: row){
                   row1.createCell(countCell).setCellValue(cell.toString());
                    countCell++;
            }
        }
        workbookNew.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }
}
