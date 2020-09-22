package com.xizi;

import com.xizi.dao.UserMapper;
import com.xizi.pojo.User;
import com.xizi.utils.MybatisUtils;
import org.apache.ibatis.session.SqlSession;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;
import java.io.FileOutputStream;
import java.util.List;


public class ExcelWrite_User {
    String PATH="D:\\IDEA\\JavaWorkSpace\\Easyexcel\\xizi-poi\\";
    @Test
    public void testWrite() throws Exception {
        //1.创建一个工作簿03
        Workbook workbook=new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet=workbook.createSheet("学生成绩统计表");
        //3.创建一个行 (1,1)
        Row row1=sheet.createRow(0);
        //4.创建单元格
        Cell cell11= row1.createCell(0);
        cell11.setCellValue("姓名");
        //（1，2）
        Cell cell12=row1.createCell(1);
        cell12.setCellValue("学号");
        //（1，3）
        Cell cell13=row1.createCell(2);
        cell13.setCellValue("学院");
        //(1,4)
        Cell cell14=row1.createCell(3);
        cell14.setCellValue("综合素质分数");
        //（1，5）
        Cell cell5=row1.createCell(4);
        cell5.setCellValue("考试成绩分数");

        // 第二行(2,1)
        for (int i = 1; i < 10; i++) {
            Row row2 = sheet.createRow(i);
            Cell name = row2.createCell(0);
            name.setCellValue("戏子"+i);
            Cell id = row2.createCell(1);
            id.setCellValue("180207012"+i);
            Cell college = row2.createCell(2);
            college.setCellValue("人工智能学院");
            Cell score1 = row2.createCell(3);
            score1.setCellValue(Math.ceil(Math.random()*100));
            Cell score2 = row2.createCell(4);
            score2.setCellValue(Math.ceil(Math.random()*100));
        }


        //生成一张表  (IO流)  03 版本就是使用xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "学生成绩统计表03.xls");
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
    }

    @Test
    public void testWrite_user() throws Exception {

        //第一步：获取sqlSession对象
        SqlSession sqlSession = MybatisUtils.getSqlSession();
        //执行SQL
        UserMapper userDao = sqlSession.getMapper(UserMapper.class);
        List<User> userList = userDao.getUserList();


        //1.创建一个工作簿03
        Workbook workbook=new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet=workbook.createSheet("用户统计表");
        //3.创建一个行 (1,1)
        Row row1=sheet.createRow(0);
        //4.创建单元格
        Cell cell11= row1.createCell(0);
        cell11.setCellValue("学号");
        //（1，2）
        Cell cell12=row1.createCell(1);
        cell12.setCellValue("用户名");
        //（1，3）
        Cell cell13=row1.createCell(2);
        cell13.setCellValue("密码");

        // 第二行(2,1)

        for (int i = 1; i <= userList.size(); i++) {
            Row row2 = sheet.createRow(i);
            Cell name = row2.createCell(0);
            name.setCellValue(userList.get(i-1).getId());
            Cell id = row2.createCell(1);
            id.setCellValue(userList.get(i-1).getName());
            Cell college = row2.createCell(2);
            college.setCellValue(userList.get(i-1).getPwd());
        }


        //生成一张表  (IO流)  03 版本就是使用xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "用户统计表03.xls");
        workbook.write(fileOutputStream);
        //关闭流
        sqlSession.close();
        fileOutputStream.close();
    }

    @Test
    public void testWrite07() throws Exception {
        //1.创建一个工作簿07
        Workbook workbook=new XSSFWorkbook();
        //2.创建一个工作表
        //2.创建一个工作表
        Sheet sheet=workbook.createSheet("学生成绩统计表");
        //3.创建一个行 (1,1)
        Row row1=sheet.createRow(0);
        //4.创建单元格
        Cell cell11= row1.createCell(0);
        cell11.setCellValue("姓名");
        //（1，2）
        Cell cell12=row1.createCell(1);
        cell12.setCellValue("学号");
        //（1，3）
        Cell cell13=row1.createCell(2);
        cell13.setCellValue("学院");
        //(1,4)
        Cell cell14=row1.createCell(3);
        cell14.setCellValue("综合素质分数");
        //（1，5）
        Cell cell5=row1.createCell(4);
        cell5.setCellValue("考试成绩分数");

        // 第二行(2,1)
        for (int i = 1; i < 10; i++) {
            Row row2 = sheet.createRow(i);
            Cell name = row2.createCell(0);
            name.setCellValue("戏子"+i);
            Cell id = row2.createCell(1);
            id.setCellValue("180207012"+i);
            Cell college = row2.createCell(2);
            college.setCellValue("人工智能学院");
            Cell score1 = row2.createCell(3);
            score1.setCellValue(Math.ceil(Math.random()*100));
            Cell score2 = row2.createCell(4);
            score2.setCellValue(Math.ceil(Math.random()*100));
        }

        //生成一张表  (IO流)  03 版本就是使用xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "学生成绩统计表07.xlsx");
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
    }


    @Test
    public void testWrite03BigData() throws Exception {
        //时间
        long begin =System.currentTimeMillis();

        //1.创建一个工作簿03
        Workbook workbook=new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet=workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum <65536 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum <10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWriteBigData03.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        long end =System.currentTimeMillis();
        System.out.println(end-begin);
    }
    @Test
    public void testWrite07BigData() throws Exception {
        //时间
        long begin =System.currentTimeMillis();

        //1.创建一个工作簿07
        Workbook workbook=new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet=workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum <65537 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum <10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWriteBigData07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        long end =System.currentTimeMillis();
        System.out.println(end-begin);
    }
    @Test
    public void testWrite07BigDataS() throws Exception {
        //时间
        long begin =System.currentTimeMillis();

        //1.创建一个工作簿07
        Workbook workbook=new SXSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet=workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum <65537 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum <10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWriteBigData07S.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //清除临时文件
        ((SXSSFWorkbook) workbook).dispose();
        long end =System.currentTimeMillis();
        System.out.println(end-begin);
    }
}

