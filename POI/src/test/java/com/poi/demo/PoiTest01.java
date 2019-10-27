package com.poi.demo;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class PoiTest01 {

    public static void main(String[] args) throws Exception {

        //1.创建工作簿  HSSFWorkbook -- 2003
        Workbook wb = new XSSFWorkbook(); //2007版本
        //2.创建表单sheet
        Sheet sheet = wb.createSheet("test01");
        //3.文件流
        FileOutputStream fos = new FileOutputStream("E:\\testFile.xlsx");
        //4.写入文件
        wb.write(fos);
        //5.关闭流
        fos.close();
    }

}