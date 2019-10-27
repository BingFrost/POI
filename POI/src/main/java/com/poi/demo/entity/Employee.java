package com.poi.demo.entity;

import com.poi.demo.annotation.ExcelAttribute;
import lombok.Data;

@Data
public class Employee {

    /**
     * 编号
     */
    @ExcelAttribute(sort = 0)
    private String userId;
    /**
     * 姓名
     */
    @ExcelAttribute(sort = 1)
    private String username;
    /**
     * 手机
     */
    @ExcelAttribute(sort = 2)
    private String mobile;
    /**
     * 最高学历
     */
    @ExcelAttribute(sort = 3)
    private String theHighestDegreeOfEducation;
    /**
     * 入职时间
     */
    @ExcelAttribute(sort = 5)
    private String timeOfEntry;
    /**
     * 生日
     */
    @ExcelAttribute(sort = 4)
    private String birthday;
    /**
     * 年龄
     */
    private String age;
    /**
     * 离职时间
     */
    @ExcelAttribute(sort = 8)
    private String resignationTime;
    /**
     * 离职类型
     */
    @ExcelAttribute(sort = 6)
    private String typeOfTurnover;
    /**
     * 申请离职原因
     */
    @ExcelAttribute(sort = 7)
    private String reasonsForLeaving;
}
