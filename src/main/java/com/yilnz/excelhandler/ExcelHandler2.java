package com.yilnz.excelhandler;

import cn.hutool.poi.excel.ExcelUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExcelHandler2 {

    public static void removeRow(Sheet sheet, int rowIndex) {
        final Row row = sheet.getRow(rowIndex);
        if (row != null) {
            sheet.removeRow(row);
        }
       /* int lastRowNum=sheet.getLastRowNum();
       if(rowIndex>=0&&rowIndex<lastRowNum){
            sheet.shiftRows(rowIndex+1,lastRowNum, -1);
        }*/
        /*if(rowIndex==lastRowNum){
            Row removingRow=sheet.getRow(rowIndex);
            if(removingRow!=null){
                sheet.removeRow(removingRow);
            }
        }*/
       /* for (int i = rowIndex; i < lastRowNum; i++) {
            sheet.getRow(rowIndex).setR
        }*/
    }

    public  void handleExcelDelFirstGroupLine(List<File> files) throws IOException {
        for (File file : files) {
            if(!(file.getName().endsWith(".xlsx") || file.getName().endsWith(".xls"))){
                continue;
            }
            final Workbook workbook = ExcelHandler.getWorkbook(file);
            final int numberOfSheets = workbook.getNumberOfSheets();
            List<Integer> deleteRowNums = new ArrayList();
            for(int i =0 ;i < numberOfSheets;i++){
                final Sheet sheet = workbook.getSheetAt(i);
                for(int j = 0;j <= sheet.getLastRowNum();j++){
                    final Row row = sheet.getRow(j);
                    Cell cell = null;
                    Cell cell1 = null;
                    if (row != null) {
                         cell = row.getCell(0);
                        cell1 = row.getCell(1);
                    }
                    if (row == null ||  (cell != null && cell1 != null && StringUtils.isBlank(cell.toString()) && StringUtils.isBlank(cell1.toString()))) {
                        if (sheet.getRow(j + 1) != null) {
                            deleteRowNums.add(j + 1);
                        }
                    }

                }

                //ExcelHandler.removeMerged(sheet);

               // sheet.removeRow(sheet.getRow(1));

               // sheet.shiftRows(1, sheet.getLastRowNum() + 1, -1);
                //removeRow(sheet, deleteRowNums.get(j));

             /*   for (int i1 = 0; i1 < deleteRowNums.size(); i1++) {
                    removeRow(sheet, deleteRowNums.get(i));
                }*/

               for (int j = deleteRowNums.size() - 1; j>= 0;j--) {
                    //Row deleteRow = sheet.getRow(deleteRowNums.get(j));
                   removeRow(sheet, deleteRowNums.get(j));
                    //sheet.removeRow(deleteRow);
                    //sheet.shiftRows(deleteRowNums.get(j), deleteRowNums.get(j), 0);
                }

              /*  for(int j = deleteRowNums.size(); i >= 0;j--){
                    sheet.removeRow();
                }*/
            }
           final File newFile = ExcelHandler.writeWorkBook(workbook, file);
            new ExcelHandler().handleExcelSeqSheets(newFile);
            newFile.delete();
        }
    }

    public static void main(String[] args) throws IOException {
        new ExcelHandler2().handleExcelDelFirstGroupLine(Arrays.asList(new File("/Users/zyl/Documents/itaojingit/excelhandler/电影话题编写-RHXC12-18-944组.xlsx")));
    }
}
