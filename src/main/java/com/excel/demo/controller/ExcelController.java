package com.excel.demo.controller;

import com.excel.demo.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * ClassName ExcelController
 * Description TODO
 * Author lyk
 * Date 2018/7/17 14:19
 * Version 1.0
 **/
@Controller
@RequestMapping("/excel")
public class ExcelController {
    @Autowired
    private ExcelService excelService;

    @RequestMapping("/file")
    public String file(){
        return "import";
    }
    @RequestMapping("/import")
    @ResponseBody
    public String dataImport(@RequestParam("fileName") MultipartFile file){
        String fileName = file.getOriginalFilename();
//        if (fileName.isEmpty()){
//
//        }
        try {
              excelService.importExcel(fileName,file);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "导入成功！";
    }
    @RequestMapping("/export")
    @ResponseBody
    public String dataExport(HttpServletResponse response){
        try {
            excelService.exportExcel(response);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

}
