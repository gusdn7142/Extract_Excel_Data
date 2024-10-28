package com.example.extract_excel_data_test.controller;

import com.example.extract_excel_data_test.model.ExcelDataInfo;
import com.example.extract_excel_data_test.service.ExcelService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@Slf4j
@RestController
@RequestMapping("/excelfile")
public class ExcelController {

    private ExcelService excelService;

    @Autowired
    public ExcelController(ExcelService excelService) {
        this.excelService = excelService;
    }

    @PostMapping("/extractdata")
    public void extractDataFromExcel(@RequestParam("excelFile") MultipartFile excelFile) throws IOException, IllegalAccessException {
        List<ExcelDataInfo> excelDataInfoList = excelService.extractDataFromV2(excelFile, ExcelDataInfo::new);
        excelDataInfoList.stream().forEach(excelDataInfo -> {log.info("excelDataInfo = {}",excelDataInfo);});
    }
}
