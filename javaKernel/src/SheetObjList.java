import java.io.*;
import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetObjList {
    private FileInputStream file;//File
    private XSSFWorkbook workbook;//Excel workbooks
    private XSSFSheet sheet;//sheet in workbook

    //Constructor
    public SheetObjList(String pathToTheFile, int sheetIndex) {
        try {
            file = new FileInputStream(new File(pathToTheFile));//It opens file.
            workbook = new XSSFWorkbook(file);//This is excels file. It appoints a workbook.
            sheet = workbook.getSheetAt(sheetIndex);//Sheet number from 0 to ...
            System.out.println("The file was opened correctly.");
        } catch (Exception e) {
            System.out.println("File : \"" + pathToTheFile + "\", do not found. " + e);
        }
    }

    public boolean isTypeEquals(String typeName, int rowIndex) throws java.io.IOException {
        final int ROW_DEVISE_TYPE_INDEX = 3;
        XSSFRow row = sheet.getRow(rowIndex);
        XSSFCell cell = row.getCell(ROW_DEVISE_TYPE_INDEX);
        String s = cell.getStringCellValue();
        file.close();
        if (s.equals(typeName)) {
            return true;
        } else {
            return false;
        }

    }

    private void setX_ByType() {

    }

}
