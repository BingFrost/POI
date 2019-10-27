package com.poi.demo.controller;

import com.poi.demo.entity.Employee;
import com.poi.demo.entity.User;
import com.poi.demo.service.UserService;
import com.poi.demo.utils.DownloadUtils;
import com.poi.demo.utils.ExcelExportUtil;
import com.poi.demo.utils.ExcelImportUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

@RestController
@RequestMapping(value = "/sys")
public class UserController {

    @Autowired
    UserService userService;

    /**
     * excel 导入
     */
    @RequestMapping(value = "/user/import", method = RequestMethod.POST)
    public String importExcel(@RequestParam(name = "file") MultipartFile attachment) throws Exception {
        //根据上传流信息创建工作簿
        Workbook workbook = WorkbookFactory.create(attachment.getInputStream());
        //获取第一个sheet
        Sheet sheet = workbook.getSheetAt(0);
        List<User> users = new ArrayList<>();
        //从第二行开始获取数据
        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            Object objs[] = new Object[row.getLastCellNum()];
            //从第二列获取数据
            for (int cellNum = 1; cellNum < row.getLastCellNum(); cellNum++) {
                Cell cell = row.getCell(cellNum);
                objs[cellNum] = getCellValue(cell);
            }
            //根据每一列构造用户对象
            User user = new User(objs);
            users.add(user);
        }

        System.out.println(users.size());
        for (User user : users) {
            System.out.println(user.toString());
        }
        return "操作成功！";
    }

    public static Object getCellValue(Cell cell) {
        //1.获取到单元格的属性类型
        CellType cellType = cell.getCellType();
        //2.根据单元格数据类型获取数据
        Object value = null;
        switch (cellType) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    //日期格式
                    value = cell.getDateCellValue();
                } else {
                    //数字
                    value = cell.getNumericCellValue();
                }
                break;
            case FORMULA: //公式
                value = cell.getCellFormula();
                break;
            default:
                break;
        }
        return value;
    }

    /**
     * 根据月份参数导出当月数据
     */
    @RequestMapping(value = "/export/{month}", method = RequestMethod.GET)
    public void export(@PathVariable(name = "month") String month) throws Exception {

        ServletRequestAttributes sra = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = sra.getResponse();

        //1.获取数据。 加百分号是方便后期使用LIKE查询
        List<Employee> list = userService.findByMonth(month + "%");
        //2.创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //3.构造sheet
        Sheet sheet = workbook.createSheet();
        //4.标题
        String[] titles = {"编号", "姓名", "手机", "最高学历", "生日", "入职时间", "离职类型", "离职原因", "离职时间"};
        Row row = sheet.createRow(0);

//        int titleIndex=0;
//        for (String title : titles) {
//            Cell cell = row.createCell(titleIndex++);
//            cell.setCellValue(title);
//        }
        AtomicInteger headersAi = new AtomicInteger(); // 从索引0开始
        for (String title : titles) {
            Cell cell = row.createCell(headersAi.getAndIncrement());
            cell.setCellValue(title);
        }
        AtomicInteger datasAi = new AtomicInteger(1);
        Cell cell = null;
        for (Employee employee : list) {
            Row dataRow = sheet.createRow(datasAi.getAndIncrement());
            //编号
            cell = dataRow.createCell(0);
            cell.setCellValue(employee.getUserId());
            //姓名
            cell = dataRow.createCell(1);
            cell.setCellValue(employee.getUsername());
            //手机
            cell = dataRow.createCell(2);
            cell.setCellValue(employee.getMobile());
            //最高学历
            cell = dataRow.createCell(3);
            cell.setCellValue(employee.getTheHighestDegreeOfEducation());
            //生日
            cell = dataRow.createCell(4);
            cell.setCellValue(employee.getBirthday());
            //入职时间
            cell = dataRow.createCell(5);
            cell.setCellValue(employee.getTimeOfEntry());
            //离职类型
            cell = dataRow.createCell(6);
            cell.setCellValue(employee.getTypeOfTurnover());
            //离职原因
            cell = dataRow.createCell(7);
            cell.setCellValue(employee.getReasonsForLeaving());
            //离职时间
            cell = dataRow.createCell(8);
            cell.setCellValue(employee.getResignationTime());
        }

        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);
        new DownloadUtils().download(os, response, month + "月份报表.xlsx");
    }

    /**
     * 模板打印
     */
    @RequestMapping(value = "/exportbytemplate/{month}", method = RequestMethod.GET)
    public void exportByTemplate(@PathVariable(name = "month") String month) throws Exception {

        // 获取response
        ServletRequestAttributes sra = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = sra.getResponse();

        //1.获取数据。 加百分号是方便后期使用LIKE查询
        List<Employee> list = userService.findByMonth(month + "%");
        //2.加载模板
        Resource resource = new ClassPathResource("excel/template.xlsx");
        FileInputStream fis = new FileInputStream(resource.getFile());
        //3.创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        //4.构造sheet
        Sheet sheet = workbook.getSheetAt(0);
        //5.抽取公共样式，放入集合中
        Row row = sheet.getRow(2);
        CellStyle styles [] = new CellStyle[row.getLastCellNum()];
        for(int i=0;i<row.getLastCellNum();i++) {
            Cell cell = row.getCell(i);
            styles[i] = cell.getCellStyle();
        }

        // 构造单元格
        Cell cell = null;
        AtomicInteger datasAi = new AtomicInteger(2);

        for (Employee employee : list) {
            Row dataRow = sheet.createRow(datasAi.getAndIncrement());
            //编号
            cell = dataRow.createCell(0);
            cell.setCellValue(employee.getUserId());
            cell.setCellStyle(styles[0]);
            //姓名
            cell = dataRow.createCell(1);
            cell.setCellValue(employee.getUsername());
            cell.setCellStyle(styles[1]);
            //手机
            cell = dataRow.createCell(2);
            cell.setCellValue(employee.getMobile());
            cell.setCellStyle(styles[2]);
            //最高学历
            cell = dataRow.createCell(3);
            cell.setCellValue(employee.getTheHighestDegreeOfEducation());
            cell.setCellStyle(styles[3]);
            //生日
            cell = dataRow.createCell(4);
            cell.setCellValue(employee.getBirthday());
            cell.setCellStyle(styles[4]);
            //入职时间
            cell = dataRow.createCell(5);
            cell.setCellValue(employee.getTimeOfEntry());
            cell.setCellStyle(styles[5]);
            //离职类型
            cell = dataRow.createCell(6);
            cell.setCellValue(employee.getTypeOfTurnover());
            cell.setCellStyle(styles[6]);
            //离职原因
            cell = dataRow.createCell(7);
            cell.setCellValue(employee.getReasonsForLeaving());
            cell.setCellStyle(styles[7]);
            //离职时间
            cell = dataRow.createCell(8);
            cell.setCellValue(employee.getResignationTime());
            cell.setCellStyle(styles[8]);
        }

        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);
        new DownloadUtils().download(os, response, month + "月份报表.xlsx");
    }

    /**
     * 使用工具类导入excel
     */
    @RequestMapping(value = "/util/import", method = RequestMethod.POST)
    public String importExcelByutils(@RequestParam(name = "file") MultipartFile file) throws Exception {
        List<User> list = new ExcelImportUtil(User.class).readExcel(file.getInputStream(), 1, 1);
        for (User user : list) {
            System.out.println(user.toString());
        }
        return "导入成功";
    }

    /**
     * 根据月份参数导出当月数据
     */
    @RequestMapping(value = "/util/export/{month}", method = RequestMethod.GET)
    public void exportByUtils(@PathVariable(name = "month") String month) throws Exception {

        //1.获取数据。 加百分号是方便后期使用LIKE查询
        List<Employee> list = userService.findByMonth(month + "%");

        //2.加载模板流数据
        Resource resource = new ClassPathResource("excel/template.xlsx");
        FileInputStream fis = new FileInputStream(resource.getFile());

        //3.获取response
        ServletRequestAttributes sra = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = sra.getResponse();

        new ExcelExportUtil(Employee.class,2,2).export(response,fis,list,month+"月人事报表.xlsx");
    }

    /**
     * 百万数据导出
     */
    @RequestMapping(value = "/exportMillion/{month}", method = RequestMethod.GET)
    public void exportMillion(@PathVariable(name = "month") String month) throws Exception {

        ServletRequestAttributes sra = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = sra.getResponse();

        //1.获取数据。 加百分号是方便后期使用LIKE查询
        List<Employee> list = userService.findByMonth(month + "%");
        //2.创建工作簿
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        //3.构造sheet
        Sheet sheet = workbook.createSheet();
        //4.标题
        String[] titles = {"编号", "姓名", "手机", "最高学历", "生日", "入职时间", "离职类型", "离职原因", "离职时间"};
        Row row = sheet.createRow(0);

        int titleIndex=0;
        for (String title : titles) {
            Cell cell = row.createCell(titleIndex++);
            cell.setCellValue(title);
        }

        int rowIndex=0;
        Cell cell = null;
        for (int i = 0; i < 100000; i++) {
            for (Employee employee : list) {
                Row dataRow = sheet.createRow(rowIndex++);
                //编号
                cell = dataRow.createCell(0);
                cell.setCellValue(employee.getUserId());
                //姓名
                cell = dataRow.createCell(1);
                cell.setCellValue(employee.getUsername());
                //手机
                cell = dataRow.createCell(2);
                cell.setCellValue(employee.getMobile());
                //最高学历
                cell = dataRow.createCell(3);
                cell.setCellValue(employee.getTheHighestDegreeOfEducation());
                //生日
                cell = dataRow.createCell(4);
                cell.setCellValue(employee.getBirthday());
                //入职时间
                cell = dataRow.createCell(5);
                cell.setCellValue(employee.getTimeOfEntry());
                //离职类型
                cell = dataRow.createCell(6);
                cell.setCellValue(employee.getTypeOfTurnover());
                //离职原因
                cell = dataRow.createCell(7);
                cell.setCellValue(employee.getReasonsForLeaving());
                //离职时间
                cell = dataRow.createCell(8);
                cell.setCellValue(employee.getResignationTime());
            }
        }

        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);
        new DownloadUtils().download(os, response, month + "月份报表.xlsx");
    }



}