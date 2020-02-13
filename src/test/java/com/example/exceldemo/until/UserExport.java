package com.example.exceldemo.until;

import com.example.exceldemo.entity.User;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * UserExport
 *
 * @author zuojie
 * @version 1.0
 * @Date 2020/2/12 0012
 * @Time 14:35
 **/
public class UserExport {
    public void userExport(String filePath, List<User> userList){
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();

        HSSFRow titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue("用户名");
        titleRow.createCell(1).setCellValue("密码");
        titleRow.createCell(2).setCellValue("年龄");
        titleRow.createCell(3).setCellValue("地址");

        for (User user : userList) {
            int lastRowNum = sheet.getLastRowNum();//获取最后一行行号

            HSSFRow row = sheet.createRow(lastRowNum+1);
            row.createCell(0).setCellValue(user.getUserName());
//            row.createCell(1).setCellValue(user.getPassword());
//            row.createCell(2).setCellValue(user.getAge());
//            row.createCell(3).setCellValue(user.getAddress());
        }

        try {
            FileOutputStream fos = new FileOutputStream(filePath);
            workbook.write(fos);
//            workbook.close();
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
