package com.poi.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class PoiTest03 {

    public static void main(String[] args) throws Exception {

        // 创建工作簿  HSSFWorkbook -- 2003
        Workbook wb = new XSSFWorkbook(); //2007版本
        // 创建表单sheet
        Sheet sheet = wb.createSheet("test");
        // 创建行对象  参数：索引（从0开始）
        Row row = sheet.createRow(5);
        // 创建单元格对象  参数：索引（从0开始）
        Cell cell = row.createCell(5);
        // 向单元格中写入内容
        cell.setCellValue("写入测试");

        //创建单元格样式对象
        CellStyle cellStyle = wb.createCellStyle();
        //设置边框
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT);//下边框
        cellStyle.setBorderTop(BorderStyle.HAIR);//上边框
        //设置字体
        Font font = wb.createFont();//创建字体对象
        font.setFontName("华文行楷");//设置字体
        font.setFontHeightInPoints((short) 28);//设置字号
        cellStyle.setFont(font);
        //设置宽高
        sheet.setColumnWidth(0, 31 * 256);//设置第一列的宽度是31个字符宽度
        row.setHeightInPoints(50);//设置行的高度是50个点
        //设置居中显示
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        //设置单元格样式
        cell.setCellStyle(cellStyle);
        //合并单元格
        CellRangeAddress region = new CellRangeAddress(0, 3, 0, 2);
        sheet.addMergedRegion(region);

        // 文件流
        FileOutputStream fos = new FileOutputStream("E:\\testFile03.xlsx");
        // 写入文件
        wb.write(fos);
        // 关闭流
        fos.close();
    }

}