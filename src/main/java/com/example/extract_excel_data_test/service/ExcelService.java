package com.example.extract_excel_data_test.service;


import com.example.extract_excel_data_test.model.ExcelDataInfo;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Supplier;

@Slf4j
@Service
public class ExcelService {

    private final DataFormatter formatter = new DataFormatter();

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

    public <T> List<T> extractDataFromV2(MultipartFile multipartFile, Supplier<T> dataSupplier) throws IOException, IllegalAccessException {  //T excelData

        //1) CREATE Workbook (==XSSFWorkbook)
        Workbook workbook = WorkbookFactory.create(multipartFile.getInputStream());

        //2) GET Sheet
        Sheet worksheet = workbook.getSheetAt(0);

        //3) EXTRACT Data
        List<T> excelDataInfoList = new ArrayList<>();

        //3-1) Extract cellValue by row in excel
        for (int rowIndex = 1; rowIndex <= worksheet.getLastRowNum(); rowIndex++) {
            //System.out.println("getPhysicalNumberOfRows : " + worksheet.getPhysicalNumberOfRows());
            //System.out.println("getLastRowNum : " + worksheet.getLastRowNum());

            Row row = worksheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            //ExcelDataInfo excelData = ExcelDataInfo.builder().build();
            T excelData = dataSupplier.get();
            Field[] fields = excelData.getClass().getDeclaredFields();

            for (int cellIndex = 0; cellIndex < fields.length; cellIndex++) {
                //3-1-1) Get ExcelDataInfo's Field
                Field field = fields[cellIndex];
                field.setAccessible(true);         //okay - get private field

                //3-1-2) Get Cell's Value
                String cellValue = formatter.formatCellValue(row.getCell(cellIndex));

                //3-1-3) Set Cell's Value In excelData
                field.set(excelData, cellValue);
            }

            excelDataInfoList.add(excelData);
        }

        return excelDataInfoList;
    }
}
