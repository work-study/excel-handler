package com.java.code.springboot.model;

import lombok.Data;
import lombok.ToString;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

/**
 * @author zouw
 * @date 19:09 2019/1/16
 */
@Data
@ToString
public class ExcelModel implements Serializable {
    private static final long serialVersionUID = -2887667350485195486L;
    private String sheetName;
    private List<String> title;
    private List<Map<String, String>> contextList;
}
