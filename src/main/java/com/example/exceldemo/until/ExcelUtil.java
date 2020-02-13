package com.example.exceldemo.until;

import com.example.exceldemo.entity.User;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * ExcelUtil
 *
 * @author zuojie
 * @version 1.0
 * @Date 2020/2/13 0013
 * @Time 9:03
 **/
public class ExcelUtil {
    public String exportExce(List<User> data){
        String result="xlsm created failed..";
        String fileName = "./src/main/resources/templates/new_file.xlsm";
        String filePath = "./src/main/resources/static/demo.xlsm";
//        Resource resource = new ClassPathResource(filePath);//用来读取resources下的文件

        List<User> list = new ArrayList<User>();
        for (int i = 0; i < 5; i++) {
            User user = new User();
            user.setUserName("姓名"+i);
            user.setPhone("123456");
            list.add(user);
        }
        try {
//            File file = resource.getFile();
            Workbook workbook = new XSSFWorkbook(OPCPackage.open("./src/main/resources/static/demo.xlsm"));

            XSSFSheet sheet1 = ((XSSFWorkbook) workbook).getSheetAt(0);
            insertData(list,sheet1);

            XSSFSheet sheet2 = ((XSSFWorkbook) workbook).getSheetAt(1);
            insertData(data,sheet2);

            FileOutputStream out = new FileOutputStream(new File(fileName));
            workbook.write(out);
            out.close();
            System.out.println("xlsm created successfully..");
            result = "xlsm created successfully..";

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }

    public void insertData(List<User> list,XSSFSheet sheet){
        int rowIndex = 1;
        for (User user : list) {
            XSSFRow row = sheet.getRow(rowIndex);
            if (null == row) {
                row = sheet.createRow(rowIndex);
            }
            XSSFCell cell0 = row.getCell(0);
            if (null == cell0) {
                cell0 = row.createCell(0);
            }
            cell0.setCellValue(user.getUserName());// 姓名

            XSSFCell cell1 = row.getCell(1);
            if (null == cell1) {
                cell1 = row.createCell(1);
            }
            cell1.setCellValue(user.getPhone());// 电话号码
            rowIndex++;
        }
    }
}
