package com.example.extract_excel_data_test.service;


import com.example.extract_excel_data_test.model.ExcelDataInfo;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
@Service
public class ExcelService {

    public List<ExcelDataInfo> extractDataFrom(MultipartFile multipartFile) throws IOException {

        //1) CREATE Workbook (==XSSFWorkbook)
        Workbook workbook = WorkbookFactory.create(multipartFile.getInputStream());

        //2) GET Sheet
        Sheet worksheet = workbook.getSheetAt(0);

        //3) EXTRACT Data
        List<ExcelDataInfo> excelDataInfoList = new ArrayList<>();

        //3-1) Extract cellValue by row in excel
        for (int i = 1; i <= worksheet.getLastRowNum(); i++) {

            Row row = worksheet.getRow(i);
            if(row == null){
                continue;
            }

            ExcelDataInfo excelData = new ExcelDataInfo();
            excelData.updateName(row.getCell(0) != null ? row.getCell(0).getStringCellValue() : "");
            excelData.updateAddress(row.getCell(1) != null ? row.getCell(1).getStringCellValue() : "");
            excelData.updatePhone(row.getCell(2) != null ? row.getCell(2).getStringCellValue() : "");
            excelData.updateEmail(row.getCell(3) != null ? row.getCell(3).getStringCellValue() : "");
            excelData.updateComment(row.getCell(4) != null ? row.getCell(4).getStringCellValue() : "");
            excelData.updateBirthDay(row.getCell(5) != null ? row.getCell(5).getStringCellValue() : "");
            excelData.updateJob(row.getCell(6) != null ? row.getCell(6).getStringCellValue() : "");

            excelDataInfoList.add(excelData);
        }

        return excelDataInfoList;
    }
}
