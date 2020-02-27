package com.lifp.test.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;

/**
 * 测试poi操作excel
 * HSSF － 提供读写Microsoft Excel格式档案的功能。（.xls）
 * XSSF － 提供读写Microsoft Excel OOXML格式档案的功能。（.xlsx）
 * 
 * @author lifp
 * @Date 2020/2/27
 */
public class PoiDemoTest {

    /**
     * 使用03版本操作大数据量文件
     * SXSSFWorkbook 适合操作大数据量文件，经过优化的，分批次写入。效率特别高
     * XSSFWorkbook: 操作大数据量文件时效率特别低
     * 
     * @Author lifp
     * @Date 2020/2/27  
     */
    @Test
    public void testBigData07() throws IOException {

        //记录开始时间
        long begin = System.currentTimeMillis();

        //创建一个SXSSFWorkbook
        Workbook workbook = new SXSSFWorkbook();

        //创建一个sheet
        Sheet sheet = workbook.createSheet();

        //xls文件最大支持65536行
        for (int rowNum = 0; rowNum < 655360; rowNum++) {
            //创建一个行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {//创建单元格
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        System.out.println("done");
        
        FileOutputStream out = new FileOutputStream("E:\\JavaProject\\poiDemo\\file\\big07.xlsx");
        workbook.write(out);
        // 操作结束，关闭文件
        out.close();

        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();

        //记录结束时间
        long end = System.currentTimeMillis();
        System.out.println((double)(end - begin)/1000);
    }
    
    /**
     * 使用03版本操作大数据量文件
     * 
     * @Author lifp 
     * @Date 2020/2/27  
     */
    @Test
    public void testBigData03() throws IOException {

        //记录开始时间
        long begin = System.currentTimeMillis();

        //创建一个SXSSFWorkbook
        //-1：关闭 auto-flushing，将所有数据存在内存中
        Workbook workbook = new HSSFWorkbook();

        //创建一个sheet
        Sheet sheet = workbook.createSheet();

        //xls文件最大支持65536行
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            //创建一个行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {//创建单元格
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        System.out.println("done");
        FileOutputStream out = new FileOutputStream("E:\\JavaProject\\poiDemo\\file\\big03.xls");
        workbook.write(out);
        // 操作结束，关闭文件
        out.close();

        //记录结束时间
        long end = System.currentTimeMillis();
        System.out.println((double)(end - begin)/1000);
    }
    
    /**
     * 测试向03版本(.xls)中写入数据
     * 
     * @Author lifp 
     * @Date 2020/2/27  
     */
    @Test
    public void writeExcel03() throws IOException {
        
        // 1. 创建一个workbook
        Workbook workbook = new HSSFWorkbook();
        
        // 2. 根据workbook创建sheet
        Sheet sheet = workbook.createSheet("会员列表");

        // 3. 根据sheet创建行row
        Row row = sheet.createRow(0);

        // 4. 根据row创建列cell
        Cell cell = row.createCell(0);

        // 5. 向列cell里边设置值
        cell.setCellValue("李钒平");
        
        // 6. 使用输出流写入到文件中
        OutputStream outputStream = new FileOutputStream("E:\\JavaProject\\poiDemo\\file\\01.xls");
        
        // 7. 把workbook中的内容通过输出流写入文件中
        workbook.write(outputStream);
        
        // 8. 关闭流
        outputStream.close();
    }

    /**
     * 测试向07版本(.xlsx)中写入数据
     *
     * @Author lifp
     * @Date 2020/2/27  
     */
    @Test
    public void writeExcel07() throws IOException {

        // 1. 创建一个workbook
        Workbook workbook = new XSSFWorkbook();

        // 2. 根据workbook创建sheet
        Sheet sheet = workbook.createSheet("测试列表");

        // 3. 根据sheet创建行row
        Row row = sheet.createRow(0);

        // 4. 根据row创建列cell
        Cell cell = row.createCell(0);

        // 5. 向列cell里边设置值
        cell.setCellValue("Lucy");

        // 6. 使用输出流写入到文件中
        OutputStream outputStream = new FileOutputStream("E:\\JavaProject\\poiDemo\\file\\07.xlsx");

        // 7. 把workbook中的内容通过输出流写入文件中
        workbook.write(outputStream);

        // 8. 关闭流
        outputStream.close();
    }
    
    /**
     * 测试读取03版本excel里面的内容
     * 读取excel中的值时需要根据单元框内值的类型选定对应的方法
     * 数字 -> getNumericCellValue
     * 字符串 -> getStringCellValue
     * 
     * 03版本仅支持65536行数据。超过之后将无法插入进数据，所以出现了07版本
     * 
     * @Author lifp 
     * @Date 2020/2/27  
     */
    @Test
    public void readExcel03() throws IOException {
        // 1. 获取读取文件的输入流
        InputStream inputStream = new FileInputStream("E:\\JavaProject\\poiDemo\\file\\03.xls");
        
        // 2. 创建workbook，需要把输入流传递进去
        Workbook workbook = new HSSFWorkbook(inputStream);
        
        // 3. 根据workbook获取sheet
        Sheet sheet = workbook.getSheetAt(0);

        // 4. 根据sheet获取到行row
        Row row = sheet.getRow(0);

        // 5. 根据行row获取到列cell
        Cell cell = row.getCell(0);

        // 6. 根据列cell读取到内容 
        String stringCellValue = cell.getStringCellValue();

        System.out.println(stringCellValue);
        
        inputStream.close();
    }

    /**
     * 测试读取07版本excel里面的内容
     * XSSFWorkbook：
     * HSSFWorkbook：效率高，专门处理大数据量
     * 
     * @Author lifp
     * @Date 2020/2/27  
     */
    @Test
    public void readExcel07() throws IOException {
        // 1. 获取读取文件的输入流
        InputStream inputStream = new FileInputStream("E:\\JavaProject\\poiDemo\\file\\07.xlsx");

        // 2. 创建workbook，需要把输入流传递进去
        Workbook workbook = new XSSFWorkbook(inputStream);

        // 3. 根据workbook获取sheet
        Sheet sheet = workbook.getSheetAt(0);

        // 4. 根据sheet获取到行row
        Row row = sheet.getRow(0);

        // 5. 根据行row获取到列cell
        Cell cell = row.getCell(0);

        // 6. 根据列cell读取到内容 
        String stringCellValue = cell.getStringCellValue();

        System.out.println(stringCellValue);

        inputStream.close();
    }
}
