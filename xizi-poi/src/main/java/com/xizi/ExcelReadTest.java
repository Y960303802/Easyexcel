package com.xizi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Date;

public class ExcelReadTest {
    String PATH="D:\\IDEA\\JavaWorkSpace\\Easyexcel\\xizi-poi";

    @Test
    public void testRead03() throws Exception {

        FileInputStream fileInputStream = new FileInputStream(PATH+"\\观众统计表.xls");
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        Cell cell1 = row.getCell(1);
        System.out.println(cell.getStringCellValue());
        System.out.println(cell1.getNumericCellValue());
        fileInputStream.close();

    }


    @Test
    public void testRead07() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(PATH+"\\观众统计表07.xlsx");
        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        Cell cell1 = row.getCell(1);
        System.out.println(cell.getStringCellValue());
        System.out.println(cell1.getNumericCellValue());
        fileInputStream.close();

    }


    @Test
    public void testRead08() throws Exception {

        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH+"\\明细表.xls");
        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        //获取标题内容
        Row rowTitle = sheet.getRow(0);
        if (rowTitle!=null){
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell!=null){
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue+"|");
                }
            }
            System.out.println();
        }
        //获取表中的内容
        //读取行数
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum <rowCount ; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData!=null){
                //读取列数
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum <cellCount ; cellNum++) {
                    System.out.print("["+(rowNum+1)+"-"+(cellNum+1)+"]");
                    Cell cell = rowData.getCell(cellNum);
                    //匹配数据类型
                    if (cell!=null){
                        int cellType = cell.getCellType();
                        String cellValue="";
                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING:
                                System.out.print("[String]");
                                cellValue=cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.print("[BOOLEAN]");
                                cellValue=String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:
                                System.out.print("[BLANK]");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print("[NUMERIC]");
                                if (HSSFDateUtil.isCellDateFormatted(cell)){
                                    System.out.print("[日期]");
                                    Date date = cell.getDateCellValue();
                                    cellValue=new DateTime(date).toString("yyyy-MM-dd");
                                }else {
                                    //不是日期格式，防止数字过长
                                    System.out.print("[转换为字符串输出]");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue=cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("[ERROR]");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }

        fileInputStream.close();
    }

    @Test
    public void testFormula() throws Exception {
        InputStream fileInputStream = new FileInputStream(PATH+"\\公式.xls");
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);

        //拿到计算公式
        FormulaEvaluator FormulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
        //输出单元格得内容
        int cellType = cell.getCellType();
        switch (cellType){
            case Cell.CELL_TYPE_FORMULA:
                String formula = cell.getCellFormula();
                System.out.println(formula);
                //计算
                CellValue evaluate=FormulaEvaluator.evaluate(cell);
                String s = evaluate.formatAsString();
                System.out.println(s);
                break;
        }

    }
}
