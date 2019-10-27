package com.poi.demo.service;

import com.poi.demo.entity.Employee;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class UserService {

    public List<Employee> findByMonth(String month) {

        List list = new ArrayList<Employee>();

        for (int i = 0; i < 10; i++) {
            Employee employee = new Employee();
            employee.setUserId(i+"");
            employee.setUsername("张"+i);
            employee.setMobile(i+i+"");
            employee.setTheHighestDegreeOfEducation("学龄"+i);
            employee.setTimeOfEntry("2019-09-19");
            employee.setBirthday("1991-12-25");
            employee.setAge("18");
            employee.setResignationTime("2099-12-31");
            employee.setTypeOfTurnover("100");
            employee.setReasonsForLeaving("哈哈");

            list.add(employee);
        }
        return list;
    }

}