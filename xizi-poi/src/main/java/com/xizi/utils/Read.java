package com.xizi.utils;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Date;

public class Read {
    public void testRead08(InputStream inputStream) throws Exception {

        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook(inputStream);
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

        inputStream.close();
    }
}
