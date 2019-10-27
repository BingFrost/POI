package com.poi.demo.entity;

import com.poi.demo.annotation.ExcelAttribute;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.text.DecimalFormat;
import java.util.Date;

@Data
@ToString
@NoArgsConstructor
public class User {

    /**
     * ID
     */
    private String id;
    /**
     * 手机号码
     */
    @ExcelAttribute(sort = 2)
    private String mobile;
    /**
     * 用户名称
     */
    @ExcelAttribute(sort = 1)
    private String username;
    /**
     * 密码
     */
    private String password;
    /**
     * 工号
     */
    @ExcelAttribute(sort = 3)
    private String workNumber;
    /**
     * 聘用形式
     */
    @ExcelAttribute(sort = 4)
    private Integer formOfEmployment;
    /**
     * 入职时间
     */
    @ExcelAttribute(sort = 5)
    private Date timeOfEntry;
    /**
     * 部门ID
     */
    @ExcelAttribute(sort = 6)
    private String departmentId;
    /**
     * 部门名称
     */
    private String departmentName;
    /**
     * 创建时间
     */
    private Date createTime;

    public User(Object[] values) {
        // 用户名 手机号 工号 聘用形式 入职时间	部门ID
        this.username = values[1].toString();
        //默认手机号excel读取为字符串会存在科学记数法问题，转化处理    #：没有则为空
        this.mobile = values[2].toString();
        this.workNumber = new DecimalFormat("#").format(values[3]).toString();
        this.formOfEmployment = ((Double) values[4]).intValue();
        this.timeOfEntry = (Date) values[5];
        this.departmentId = values[6].toString();
    }

}