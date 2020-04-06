import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

public class ExcellMethods {
    private String filePath;
    private CellList inputlist;
    public ExcellMethods(String filePath, CellList inputlist){
        this.filePath = filePath;
        this.inputlist =inputlist;
    }

    public void ReadandLoadExcell(){
        try {
            FileInputStream file = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet =workbook.getSheetAt(0); // 엑셀파일의 시트 불러옴
            int rowIndex =0; // 행
            int MaxrowIndex = sheet.getPhysicalNumberOfRows();
            for(rowIndex = 2; rowIndex < MaxrowIndex; rowIndex++) {
                XSSFRow row = sheet.getRow(rowIndex);
                if(row != null) {
                    String name = row.getCell(0).getStringCellValue();
                    double ratio =  row.getCell(1).getNumericCellValue();
                    inputlist.pushCell(name,ratio);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public void WriteSortResult(){
        ArrayList<Cell> list = inputlist.getCellList();
        try {
            FileInputStream fileIn = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
            XSSFSheet sheet = workbook.getSheetAt(0); // 엑셀파일의 시트 불러옴

            for (int rowIndex = 2; rowIndex < list.size()+2; rowIndex++) {
                XSSFRow row = sheet.getRow(rowIndex);
                String inputname = list.get(rowIndex-2).getname();
                double inputRatio = list.get(rowIndex-2).getRatio();
                row.getCell(0).setCellValue(inputname);
                row.getCell(1).setCellValue(inputRatio);
                row.getCell(2).setCellValue(inputRatio+"%");
            }
            FileOutputStream fileOut = new FileOutputStream(filePath);
            workbook.write(fileOut);
            fileIn.close();
            fileOut.close();
        }catch (Exception e) {
            e.printStackTrace();
        }
    }
    public void PointYellow(String inputName){
        try {
            FileInputStream file = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet =workbook.getSheetAt(0); // 엑셀파일의 시트 불러옴
            int rowIndex =0; // 행
            int MaxrowIndex = sheet.getPhysicalNumberOfRows();
            for(rowIndex = 2; rowIndex < MaxrowIndex; rowIndex++) {
                XSSFRow row = sheet.getRow(rowIndex);
                if(row != null) {
                    String name = row.getCell(0).getStringCellValue();
                    if(inputName.equals(name)){
                        CellStyle cellStyle = workbook.createCellStyle();
                        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
                        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
                        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
                        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
                        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
                        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
                        row.getCell(0).setCellStyle(cellStyle);
                        row.getCell(1).setCellStyle(cellStyle);
                        row.getCell(2).setCellStyle(cellStyle);
                    }
                }
            }

            FileOutputStream fileOut = new FileOutputStream(filePath);
            workbook.write(fileOut);
            file.close();
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
