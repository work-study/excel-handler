package com.java.code.springboot.utils;


import com.java.code.springboot.model.ExcelModel;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtils {

    private static final String excel2003L = ".xls";
    /**
     * 2007+ 版本的excel
     */
    private static final String excel2007U = ".xlsx";

    /**
     * 导出Excel
     *
     * @param sheetName sheet名称
     * @return
     */
    public static XSSFWorkbook getHSSFWorkbook(String sheetName, ExcelModel excelModel) {
        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        XSSFWorkbook wb = new XSSFWorkbook();

        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        XSSFSheet sheet = wb.createSheet(sheetName);

        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        XSSFRow row = sheet.createRow(0);

        // 第四步，创建单元格，并设置值表头 设置表头居中
        XSSFCellStyle style = wb.createCellStyle();
        // 创建一个居中格式
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        //声明列对象
        XSSFCell cell = null;
        String[] title = excelModel.getTitle().toArray(new String[1]);
        //创建标题
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }
        int i = 0;
        for (Map<String, String> map : excelModel.getContextList()) {
            if (map.size() == 0) {
                continue;
            }
            row = sheet.createRow(++i);
            int j = 0;
            for (String cellStr : map.values()) {
                row.createCell(j++).setCellValue(cellStr);
            }
        }
        return wb;
    }


    /**
     * 将流中的Excel数据转成List<Map>(读取Excel)
     *
     * @param in       输入流
     * @param fileName 文件名（判断Excel版本）
     * @return
     * @throws Exception
     */
    public static List<ExcelModel> readExcel(InputStream in, String fileName) throws Exception {
        // 根据文件名来创建Excel工作薄
        Workbook work = getWorkbook(in, fileName);
        if (null == work) {
            throw new Exception("创建Excel工作薄为空！");
        }
        List<ExcelModel> excelModelList = new ArrayList<>();

        // 遍历Excel中所有的sheet
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            ExcelModel excelModel = new ExcelModel();
            // 返回数据
            List<Map<String, String>> contextList = new ArrayList<>();
            Sheet sheet = work.getSheetAt(i);
            if (sheet == null) {
                continue;
            }
            excelModel.setSheetName(sheet.getSheetName());
            // 取第一行标题
            Row row = sheet.getRow(0);
            List<String> title = new ArrayList<>();
            if (row != null) {
                for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
                    title.add((String) getCellValue(row.getCell(y)));
                }
            }
            if (CollectionUtils.isEmpty(title)) {
                continue;
            }
            excelModel.setTitle(title);
            // 遍历当前sheet中的所有行
            for (int j = 1; j < sheet.getLastRowNum() + 1; j++) {
                row = sheet.getRow(j);
                if (row == null) {
                    continue;
                }
                Map<String, String> contextMap = new LinkedHashMap<>();
                // 遍历所有的列
                for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
                    System.out.println("y:"+y);
                    contextMap.put(title.get(y), (String) getCellValue(row.getCell(y)));
                }
                contextList.add(contextMap);
            }
            excelModel.setContextList(contextList);
            excelModelList.add(excelModel);
        }

        return excelModelList;
    }

    /**
     * 描述：根据文件后缀，自适应上传文件的版本
     *
     * @param inStr ,fileName
     * @return
     * @throws Exception
     */
    private static Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (excel2003L.equals(fileType)) {
            wb = new HSSFWorkbook(inStr);
        } else if (excel2007U.equals(fileType)) {
            wb = new XSSFWorkbook(inStr);
        } else {
            throw new Exception("解析的文件格式有误！");
        }
        return wb;
    }

    /**
     * 描述：对表格中数值进行格式化
     *
     * @param cell
     * @return
     */
    private static Object getCellValue(Cell cell) {
        Object value = null;
        DecimalFormat df = new DecimalFormat("0");
        SimpleDateFormat sdf = new SimpleDateFormat("yyy-MM-dd");
        DecimalFormat df2 = new DecimalFormat("0");

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if ("General".equals(cell.getCellStyle().getDataFormatString())) {
                    value = df.format(cell.getNumericCellValue());
                } else if ("m/d/yy".equals(cell.getCellStyle().getDataFormatString())) {
                    value = sdf.format(cell.getDateCellValue());
                } else {
                    value = df2.format(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_BLANK:
                value = "";
                break;
            default:
                break;
        }
        return value;
    }


}
