
package com.test;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIController {
    /**
     * 这种方式效率比较低并且特别占用内存，数据量越大越明显
     *
     * @param args
     * @throws FileNotFoundException
     * @throws InvalidFormatException
     */
    public static void main(String[] args) throws FileNotFoundException, InvalidFormatException {
        long startTime = System.currentTimeMillis();
        BufferedOutputStream outPutStream = null;
        XSSFWorkbook workbook = null;
        FileInputStream inputStream = null;
        String filePath = "E:\\txt\\666.xlsx";
        try {
            workbook = getWorkBook(filePath);
            XSSFSheet sheet = workbook.getSheetAt(0);
            for (int i = 0; i < 50; i++) {
                for (int z = 0; z < 10000; z++) {
                    XSSFRow row = sheet.createRow(i * 10000 + z);
                    for (int j = 0; j < 10; j++) {
                        row.createCell(j).setCellValue("你好：" + j);
                    }
                }
                //每次要获取新的文件流对象，避免将之前写入的数据覆盖掉
                outPutStream = new BufferedOutputStream(new FileOutputStream(filePath));
                workbook.write(outPutStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (outPutStream != null) {
                try {
                    outPutStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        long endTime = System.currentTimeMillis();
        System.out.println(endTime - startTime);
    }

    /**
     * 先创建一个XSSFWorkbook对象
     *
     * @param filePath
     * @return
     */
    public static XSSFWorkbook getWorkBook(String filePath) {
        XSSFWorkbook workbook = null;
        try {
            File fileXlsxPath = new File(filePath);
            BufferedOutputStream outPutStream = new BufferedOutputStream(new FileOutputStream(fileXlsxPath));
            workbook = new XSSFWorkbook();
            workbook.createSheet("测试");
            workbook.write(outPutStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return workbook;
    }

}