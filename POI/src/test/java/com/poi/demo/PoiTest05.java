package com.poi.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

/**
 * 读取excel并解析
 */
public class PoiTest05 {

    public static void main(String[] args) throws Exception {

        //1.创建workbook工作簿
        Workbook wb = new XSSFWorkbook("E:\\demo.xlsx");
        //2.获取sheet 从0开始
        Sheet sheet = wb.getSheetAt(0);
        int totalRowNum = sheet.getLastRowNum();

        Row row = null;
        Cell cell = null;

        //循环所有行
        for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
            row = sheet.getRow(rowNum);
            StringBuilder sb = new StringBuilder();
            //循环每行中的所有单元格
            for (int cellNum = 2; cellNum < row.getLastCellNum(); cellNum++) {
                cell = row.getCell(cellNum);
                sb.append(getValue(cell)).append("--");
            }
            System.out.println(sb.toString());
        }

    }

    //获取数据
    private static Object getValue(Cell cell) {
        Object value = null;
        switch (cell.getCellType()) {
            case STRING: //字符串类型
                value = cell.getStringCellValue();
                break;
            case BOOLEAN: //boolean类型
                value = cell.getBooleanCellValue();
                break;
            case NUMERIC: //数字类型（包含日期和普通数字）
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue(); // 日期
                } else {
                    value = cell.getNumericCellValue(); // 数值
                }
                break;
            case FORMULA: // 公式类型
                value = cell.getCellFormula();
                break;
            default:
                break;
        }
        return value;
    }

}