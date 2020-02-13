//package com.example.exceldemo;
//
//import com.example.exceldemo.entity.User;
//import com.example.exceldemo.until.UserExport;
//import org.junit.jupiter.api.Test;
//import org.springframework.boot.test.context.SpringBootTest;
//
//import java.util.ArrayList;
//import java.util.List;
//
//@SpringBootTest
//class ExcelDemoApplicationTests {
//
//    @Test
//    public void contextLoads() {
//        List<User> userList = new ArrayList<User>();
//        for (int i = 0; i < 10; i++) {
//            User user = new User();
//            user.setUserName("sa"+i);
//            user.setPassword("upsd"+i);
//            user.setAge(21+i);
//            user.setAddress("重庆市");
//
//            userList.add(user);
//        }
//
//        new UserExport().userExport("D://userList.xls", userList);
//    }
//
//}
