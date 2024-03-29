package com.excel.controller;

import com.alibaba.fastjson.JSON;
import com.excel.domain.BusClick;
import com.excel.service.ExcelService;
import com.excel.util.ExcelUtils;
import io.swagger.annotations.Api;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/excel")
@Api(value = "excel导入导出", tags = "excel导入导出", description = "excel导入导出")
public class ExcelController {

    @Autowired
    ExcelService excelService;

    @RequestMapping(value = "/exportExcel", method = RequestMethod.GET)
    public void exportExcel() {
        List<BusClick> resultList = excelService.getBusClick();

        long t1 = System.currentTimeMillis();
        ExcelUtils.writeExcelByAnnotation(resultList, BusClick.class, ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getResponse());
        long t2 = System.currentTimeMillis();
        System.out.println(String.format("write over! cost:%sms", (t2 - t1)));
    }

    @PostMapping(value = "/readExcel")
    public void readExcel(MultipartFile file){

        long t1 = System.currentTimeMillis();
        List<BusClick> list = ExcelUtils.readExcelObject(BusClick.class, file);
        long t2 = System.currentTimeMillis();
        System.out.println(String.format("read over! cost:%sms", (t2 - t1)));
        list.forEach(
                b -> System.out.println(JSON.toJSONString(b))
        );
    }

    @PostMapping(value = "/readMap")
    public void readMap(MultipartFile file) throws IOException {

        long t1 = System.currentTimeMillis();
        List<Map<String, String>> list = ExcelUtils.readExcelMap(file.getInputStream(), file.getOriginalFilename(), 0);
        long t2 = System.currentTimeMillis();
        System.out.println(String.format("read over! cost:%sms", (t2 - t1)));
        list.forEach(
                b -> System.out.println(JSON.toJSONString(b))
        );
    }
}
