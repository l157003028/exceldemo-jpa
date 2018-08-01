package com.excel.demo.service.impl;

import com.excel.demo.dao.UserDao;
import com.excel.demo.entity.User;
import com.excel.demo.service.ExcelService;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

/**
 * ClassName ExcelServiceImpl
 * Description TODO
 * Author lyk
 * Date 2018/7/17 15:42
 * Version 1.0
 **/
@Service
@Transactional
public class ExcelServiceImpl implements ExcelService {

    @Autowired
    private UserDao userDao;

    @Override
    public boolean importExcel(String fileName, MultipartFile file) throws Exception {
        boolean notNull = false;
        if (!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$")) {
            throw new Exception("上传文件格式错误！");
        }
        boolean isExcel2003 = true;
        if (fileName.matches("^.+\\.(?i)(xlsx)$")) {
            isExcel2003 = false;
        }
        InputStream is = file.getInputStream();
        Workbook wb = null;
        if (isExcel2003) {
            wb = new HSSFWorkbook(is);
        } else {
            wb = new XSSFWorkbook(is);
        }
        Sheet sheet = wb.getSheetAt(0);
        if (sheet != null) {
            notNull = true;
        }
        // User user = null;
        List<User> userList = new ArrayList<User>();

//        System.out.println(sheet.getFirstRowNum());
//        System.out.println(sheet.getLastRowNum());

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            User user = new User();
            if (row == null)
                continue;
            if (row.getCell(0).getCellType() != 1) {
                throw new Exception("第" + r + 1 + "行导入失败！");
            }
            String name = row.getCell(0).getStringCellValue();
            if (name == null || name.isEmpty()) {
                throw new Exception("第" + r + 1 + "行导入失败！");
            }
            row.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
            String phone = row.getCell(1).getStringCellValue();
            if (phone == null || phone.isEmpty()) {
                throw new Exception("第" + r + 1 + "行导入失败！");
            }
            String address = row.getCell(2).getStringCellValue();
            if (address == null) {
                throw new Exception("第" + r + 1 + "行导入失败！");
            }
            String email = row.getCell(3).getStringCellValue();
            if (email == null) {
                throw new Exception("第" + r + 1 + "行导入失败！");
            }
            user.setName(name);
            user.setPhone(phone);
            user.setAddress(address);
            user.setEmail(email);
            userList.add(user);
        }
        for (User users : userList) {
            userDao.save(users);
            List<User> usersList = userDao.findAll();
            if (usersList.isEmpty()) {
                return false;
            }
        }
        System.out.println("导入文件：" + fileName);
        return notNull;

    }

    @Override
    public void exportExcel(HttpServletResponse response) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        String sheetName = "用户信息表";
        HSSFSheet sheet = workbook.createSheet(sheetName);
        String fileName = sheetName.concat(".xls");
        System.out.println(fileName);

        int rowNum = 1;
        String[] header = {"姓名", "电话", "地址", "邮箱"};
        HSSFRow rows = sheet.createRow(0);
        for (int i = 0; i < header.length; i++) {
            HSSFCell cell = rows.createCell(i);
            HSSFRichTextString textString = new HSSFRichTextString(header[i]);
            cell.setCellValue(textString);
        }
        List<User> userList = userDao.findAll();
        for (User user : userList) {
            HSSFRow row = sheet.createRow(rowNum);
            row.createCell(0).setCellValue(user.getName());
            row.createCell(1).setCellValue(user.getPhone());
            row.createCell(2).setCellValue(user.getAddress());
            row.createCell(3).setCellValue(user.getEmail());
            rowNum++;
        }
        response.setContentType("application/octet-stream");
        response.setHeader("content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
        response.flushBuffer();
        workbook.write(response.getOutputStream());

        System.out.println("导出文件：" + fileName);
    }
}
