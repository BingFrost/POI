package com.poi.demo.handler;

import com.poi.demo.entity.PoiEntity;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * 自定义事件处理器：处理每一行数据读取
 */
public class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    private PoiEntity entity;

    /**
     * 当开始解析某一行时触发
     *
     * @param rowNum 行的索引
     */
    @Override
    public void startRow(int rowNum) {

        // 1.实例化对象
        if (rowNum > 0) {
            entity = new PoiEntity();
        }
    }

    /**
     * 当结束解析某一行时触发
     *
     * @param rowNum 行的索引
     */
    @Override
    public void endRow(int rowNum) {
        if (rowNum > 0) {
            // 3.使用对象进行业务操作
            System.out.println(entity.toString());
        }
    }

    /**
     * 对行中的每一个表格进行处理
     *
     * @param cellReference  单元格名称（列名号） A、B、C、D。。。
     * @param formattedValue 数据
     * @param comment        批注
     */
    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {

        if (entity != null) {
            switch (cellReference.substring(0, 1)) {
                case "A":
                    entity.setId(formattedValue);
                    break;
                case "B":
                    entity.setBreast(formattedValue);
                    break;
                case "C":
                    entity.setAdipocytes(formattedValue);
                    break;
                case "D":
                    entity.setNegative(formattedValue);
                    break;
                case "E":
                    entity.setStaining(formattedValue);
                    break;
                case "F":
                    entity.setSupportive(formattedValue);
                    break;
                default:
                    break;
            }
        }
    }

}