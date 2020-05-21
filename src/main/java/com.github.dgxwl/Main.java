package com.github.dgxwl;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        try (XSSFWorkbook workbook = new XSSFWorkbook("C:/Users/Administrator/Desktop/有删行.xlsx");
             FileOutputStream fos = new FileOutputStream("C:/Users/Administrator/Desktop/行补全.xlsx")) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            //获取列数
            int columnNum = sheet.getRow(0).getPhysicalNumberOfCells();
            //获得总行数
            int rowNum = sheet.getPhysicalNumberOfRows();

            if (columnNum < 1) {
                //TODO 提示无数据
                return ;
            }
            if (rowNum < 1) {
                //TODO 提示无数据
                return ;
            }

            XSSFRow titleRow = sheet.getRow(0);

            //创建新文件
            XSSFWorkbook newWorkBook = new XSSFWorkbook();
            newWorkBook.createSheet();
            XSSFSheet newSheetAt = newWorkBook.getSheetAt(0);
            newSheetAt.createRow(0);

            //复制表头
            for (int i = 0; i < columnNum; i++) {
                XSSFCell cell = titleRow.getCell(i);
                XSSFRow newSheetRow1 = newSheetAt.getRow(0);
                newSheetRow1.createCell(i);
                XSSFCell newCell = newSheetRow1.getCell(i);
                newCell.setCellValue(cell.getStringCellValue());
            }

            //拿到开始的序号
            XSSFRow r1 = sheet.getRow(1);
            XSSFCell c1 = r1.getCell(0);
            int startSeq = (int) c1.getNumericCellValue();  //开始序号

            //遍历原表的行
            for (int i = 1; i < rowNum; i++) {
                //拿原表每行第一列的序号
                XSSFRow row = sheet.getRow(i);
                XSSFCell cell = row.getCell(0);
                int rowSeq = (int) cell.getNumericCellValue();

                //序号对不上, 有空缺
                while (startSeq < rowSeq) {
                    //新文件里创建空行
                    XSSFRow newRow = newSheetAt.createRow(startSeq);
                    XSSFCell newCell = newRow.createCell(0);
                    newCell.setCellValue(startSeq);
                    startSeq++;
                }

                //把原有行拷贝到新文件中
                XSSFRow newRow = newSheetAt.createRow(startSeq);
                for (int j = 0; j < columnNum; j++) {
                    XSSFCell newCell = newRow.createCell(j);
                    XSSFCell cj = row.getCell(j);
                    newCell.setCellType(cj.getCellType());

                    switch (cell.getCellType()) {
                        case STRING:
                            newCell.setCellValue(cj.getStringCellValue());
                            break;
                        case NUMERIC:
                            newCell.setCellValue(cj.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            newCell.setCellValue(cj.getBooleanCellValue());
                            break;
                        case FORMULA:
                            newCell.setCellValue(cj.getCellFormula());
                            break;
                        case ERROR:
                            newCell.setCellValue(cj.getErrorCellValue());
                            break;
                        default:
                            break;
                    }
                }

                startSeq++;
            }

            newWorkBook.write(fos);
            newWorkBook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
