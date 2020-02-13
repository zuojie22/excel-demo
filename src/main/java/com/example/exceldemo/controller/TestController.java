package com.example.exceldemo.controller;


import com.example.exceldemo.entity.User;
import com.example.exceldemo.until.ExcelUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * TestController
 *
 * @author zuojie
 * @version 1.0
 * @Date 2020/2/12 0012
 * @Time 22:04
 **/
@RestController
public class TestController {

    @GetMapping("/test")
    public String test() {
        ExcelUtil excelUtil = new ExcelUtil();

        List<User> list = new ArrayList<User>();
        User user1 = new User();
        user1.setUserName("姓名3");
        user1.setPhone("123456");

        User user2 = new User();
        user2.setUserName("姓名4");
        user2.setPhone("123456");

        list.add(user1);
        list.add(user2);

        String result = excelUtil.exportExce(list);
        return result;
    }
}
