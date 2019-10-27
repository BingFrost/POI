package com.poi.demo;

import com.poi.demo.handler.SheetHandler;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;

/**
 * 使用事件模型解析百万数据excel报表
 */
public class MillionPoiTest {

    public static void main(String[] args) throws Exception {

        //1.根据Excel获取OPCPackage对象 （即将excel以压缩包的形式打开）
        String path = "E:\\millionDemo.xlsx";
        OPCPackage opcPackage = OPCPackage.open(path, PackageAccess.READ);
        try {
            // 2.创建XSSFReader对象
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            //3.获取SharedStringsTable对象
            SharedStringsTable sharedStringsTable = xssfReader.getSharedStringsTable();
            // 4.获取StylesTable对象
            StylesTable stylesTable = xssfReader.getStylesTable();
            // 5.创建Sax的XmlReader对象
            XMLReader xmlReader = XMLReaderFactory.createXMLReader();
            // 6.设置处理器
            XSSFSheetXMLHandler xssfSheetXMLHandler = new XSSFSheetXMLHandler(stylesTable, sharedStringsTable, new SheetHandler(), false);
            xmlReader.setContentHandler(xssfSheetXMLHandler);
            // 7.逐行读取
            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();

            while (sheets.hasNext()) {
                InputStream sheetstream = sheets.next(); // 每一个sheet的流数据
                InputSource sheetSource = new InputSource(sheetstream);
                try {
                    xmlReader.parse(sheetSource);
                } finally {
                    sheetstream.close();
                }
            }
        } finally {
            opcPackage.close();
        }
    }

}