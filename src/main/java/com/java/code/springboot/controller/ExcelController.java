package com.java.code.springboot.controller;

import com.java.code.springboot.model.ExcelModel;
import com.java.code.springboot.utils.ExcelUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.*;

/**
 * @author zouw
 * @date 16:58 2019/1/16
 */
@Controller
public class ExcelController {

    @RequestMapping(value = "/excel",method = RequestMethod.POST)
    @ResponseBody
    public String excel(String cellListStr, String resultCell, @RequestParam(defaultValue = "2") Integer startNum, @RequestParam("file") MultipartFile file, HttpServletResponse response) throws Exception {
        if (file.isEmpty()) {
            return "上传失败，请选择文件";
        }
        List<String> cellListTemp = Arrays.asList(cellListStr.split(","));
        InputStream fileInputStream = file.getInputStream();
        List<ExcelModel> excelModelList = ExcelUtils.readExcel(fileInputStream, file.getOriginalFilename());
        ExcelModel excelModel = excelModelList.get(0);
        List<String> title = excelModel.getTitle();
        int cellNum = title.size();
        List<String> cellList = new ArrayList<>();
        for (String s : cellListTemp) {
            try {
                Integer value = Integer.valueOf(s);
                cellList.add(title.get(value));
            } catch (NumberFormatException e) {
                cellList.add(s);
                continue;
            }
        }
        try {
            Integer value = Integer.valueOf(resultCell);
            resultCell = title.get(value);
        } catch (NumberFormatException e) {
           //
        }
        Map<String, Integer> numMap = new HashMap<>(cellNum);
        Integer i = 0;
        int j = 1;
        for (Map<String, String> map : excelModel.getContextList()) {
            if (j++ < startNum) {
                continue;
            }
            if (map.size()<=1){
                continue;
            }
            String key = null;
            for (String var1 : cellList) {
                key += map.get(var1);
            }
            Integer index = numMap.get(key);
            if (index == null) {
                String var2 = String.valueOf(++i);
                map.put(resultCell, var2);
                numMap.put(key,Integer.valueOf(var2));
            } else {
                map.put(resultCell, String.valueOf(index));
            }
        }
        String fileName="导出的excel";
        response.setHeader("Content-type","application/vnd.ms-excel");
        // 解决导出文件名中文乱码
        response.setCharacterEncoding("UTF-8");
        response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes("UTF-8"),"ISO-8859-1")+".xlsx");
        // 模板导出Excel
        XSSFWorkbook hssfWorkbook = ExcelUtils.getHSSFWorkbook("sheet1", excelModel);
        hssfWorkbook.write(response.getOutputStream());
        return "success";
    }


    @RequestMapping("/index")
    public String index() {
        return "excelPage";
    }

}
