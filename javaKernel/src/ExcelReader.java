import java.io.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.logging.Level;
import java.util.logging.Logger;

public class ExcelReader {
    private static Logger logger = Logger.getLogger(ExcelFile.class.getName()); //Це для ведення логів
    private XSSFWorkbook workbook;//Excel workbooks
    private XSSFSheet sheet;//sheet in workbook
    private FileInputStream excelFileInput;//File input
    private String pathToTheFile;
    private int sheetIndex;

    ExcelReader(String pathToTheFile, int sheetIndex) {
        this.pathToTheFile = pathToTheFile;
        this.sheetIndex = sheetIndex;
    }

    private void initializer() {
        try {
            excelFileInput = new FileInputStream(new File(this.pathToTheFile));//It opens file.
            workbook = new XSSFWorkbook(excelFileInput);//This is excels file. It appoints a workbook.
            sheet = workbook.getSheetAt(this.sheetIndex);//Sheet number from 0 to ...
        } catch (java.io.FileNotFoundException e1) {
            logger.log(Level.WARNING, "Не вдалося знайти файл." + e1);
        } catch (java.io.IOException e2) {
            logger.log(Level.WARNING, "Помилка вводу виводу." + e2);
        } catch (Exception e3) {
            logger.log(Level.WARNING, "Не вдалося створити FileInput." + e3);
        }
    }

    public String getStringValue(int rowIndex, int cellIndex) {
        try {
            initializer();
            XSSFRow row = sheet.getRow(rowIndex);
            XSSFCell cell = row.getCell(cellIndex);
            if (isTypeString(rowIndex, cellIndex)) {
                String s = cell.getStringCellValue();
                excelFileInput.close();
                return s;
            } else {
                excelFileInput.close();
                return null;
            }
        } catch (java.io.IOException e) {
            logger.log(Level.WARNING, "Не вдалося закрити FileInput." + e);
            return null;
        }
    }

    public double getNumericValue(int rowIndex, int cellIndex) {
        try {
            initializer();
            XSSFRow row = sheet.getRow(rowIndex);
            XSSFCell cell = row.getCell(cellIndex);
            if (isTypeNumeric(rowIndex, cellIndex)) {
                Double v = cell.getNumericCellValue();
                excelFileInput.close();
                return v;
            } else {
                excelFileInput.close();
                return 0.0;
            }
        } catch (java.io.IOException e) {
            logger.log(Level.WARNING, "Не вдалося закрити FileInput." + e);
            return 0.0;
        }
    }

    public boolean isTypeString(int rowIndex, int cellIndex) {
        XSSFRow row = sheet.getRow(rowIndex);
        XSSFCell cell = row.getCell(cellIndex);
        switch (cell.getCellTypeEnum()) {//Detecting data type in cells.
            case STRING:
                return true;
            case NUMERIC:
                return false;
            case BOOLEAN:
                return false;
            case FORMULA:
                return false;
            default:
                return false;
        }
    }

    public boolean isTypeNumeric(int rowIndex, int cellIndex) {
        XSSFRow row = sheet.getRow(rowIndex);
        XSSFCell cell = row.getCell(cellIndex);
        switch (cell.getCellTypeEnum()) {//Detecting data type in cells.
            case STRING:
                return false;
            case NUMERIC:
                return true;
            case BOOLEAN:
                return false;
            case FORMULA:
                return false;
            default:
                return false;
        }
    }
}
