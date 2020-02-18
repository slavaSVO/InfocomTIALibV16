import java.io.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.logging.Level;
import java.util.logging.Logger;

//Це клас для роботи із таблицею Excel він створює в собі два підкласи, один для читання інший для запису.
public class ExcelFile {
    private static Logger logger = Logger.getLogger(ExcelFile.class.getName()); //Це для ведення логів
    private ExcelInputFile inputFile;//Клас який зчитує дані із таблиці
    public ExcelOutputFile outputFile;//Клас який записує дані до таблиці

    public ExcelFile(String pathToTheFile, int sheetIndex, String sheetName) {
        inputFile = new ExcelInputFile(pathToTheFile, sheetIndex);
        outputFile = new ExcelOutputFile(inputFile.workbook, inputFile.sheet, pathToTheFile);
    }

//Функція на вхід якої потрібно подати номер рядка і тип обєкту, а вона говорить чи це такий тип.
    public boolean isTypeEquals(String typeName, int rowIndex) throws java.io.IOException {
        return inputFile.isTypeEquals(typeName, rowIndex);
    }

    public void setParametersByObjectType(String typeName, int rowIndex) {
        if (typeName.equals("XV")) {//Параметри клапана
            outputFile.setParametersByRowIndex(rowIndex, 10);//
            outputFile.setParametersByRowIndex(rowIndex, 11);//
            outputFile.setParametersByRowIndex(rowIndex, 16);//
        }
        if (typeName.equals("MR")) {//Параметри реверсивного двигуна
            outputFile.setParametersByRowIndex(rowIndex, 7);//
            outputFile.setParametersByRowIndex(rowIndex, 8);//
            outputFile.setParametersByRowIndex(rowIndex, 9);//
            outputFile.setParametersByRowIndex(rowIndex, 14);//
            outputFile.setParametersByRowIndex(rowIndex, 15);//
        }
        if (typeName.equals("M")) {//Параметри двигуна прямого пуску
            outputFile.setParametersByRowIndex(rowIndex, 7);//
            outputFile.setParametersByRowIndex(rowIndex, 8);//
            outputFile.setParametersByRowIndex(rowIndex, 14);//
        }
        if (typeName.equals("S")) {//Параметри сенсора
            outputFile.setParametersByRowIndex(rowIndex, 12);//
        }
        if (typeName.equals("A")) {//Параметри аналогового датчика
            outputFile.setParametersByRowIndex(rowIndex, 13);//
        }
    }

    public void save (){
        outputFile.save();
    }


    //-----------------------------------------------------------------------------------------------------------------
    private static class ExcelInputFile {
        private XSSFWorkbook workbook;//Excel workbooks
        private XSSFSheet sheet;//sheet in workbook
        private FileInputStream excelFileInput;//File input

        ExcelInputFile(String pathToTheFile, int sheetIndex) {
            try {
                excelFileInput = new FileInputStream(new File(pathToTheFile));//It opens file.
                workbook = new XSSFWorkbook(excelFileInput);//This is excels file. It appoints a workbook.
                sheet = workbook.getSheetAt(sheetIndex);//Sheet number from 0 to ...
            } catch (java.io.FileNotFoundException e1) {
                logger.log(Level.WARNING, "Не вдалося знайти файл." + e1);
            } catch (java.io.IOException e2) {
                logger.log(Level.WARNING, "Помилка вводу виводу." + e2);
            } catch (Exception e3) {
                logger.log(Level.WARNING, "Не вдалося створити FileInput." + e3);
            }
        }

        public boolean isTypeEquals(String typeName, int rowIndex) throws java.io.IOException {
            final int ROW_DEVISE_TYPE_INDEX = 2;
            XSSFRow row = sheet.getRow(rowIndex);
            XSSFCell cell = row.getCell(ROW_DEVISE_TYPE_INDEX);
            String s = cell.getStringCellValue();
            if (s.equals(typeName)) {
                this.excelFileInput.close();
                return true;
            } else {
                this.excelFileInput.close();
                return false;
            }
        }
    }

    //-------------------------------------------------------------------------------------------------------------
    private static class ExcelOutputFile {
        private XSSFWorkbook workbook;//Excel workbooks
        private XSSFSheet sheet;//sheet in workbook
        private FileOutputStream excelFileOut;

        ExcelOutputFile(XSSFWorkbook wb, XSSFSheet st, String pathToTheFile) {
            try {
                workbook = wb;
                sheet = st;
                excelFileOut = new FileOutputStream(pathToTheFile);
            } catch (Exception e) {
                logger.log(Level.WARNING, "Не вдалося створити OutputFile." + e);
            }
        }

        public void setParametersByRowIndex(int rowIndex, int cellIndex) {
            try {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.createCell(cellIndex);
                cell.setCellValue("X");
            } catch (Exception e) {
                logger.log(Level.WARNING, "Не записати Х в OutputFile." + e);
            }
        }

        public void save() {
            try {
                workbook.write(excelFileOut);
                excelFileOut.close();
            } catch (Exception e) {
                logger.log(Level.WARNING, "Не вдалося зберегти і закрити OutputFile." + e);
            }
        }
    }
}
