package com.paul.paulexcel.poi.bean;

import com.alibaba.excel.annotation.ExcelProperty;
import com.paul.paulexcel.poi.service.impl.DealSBExcelServiceImpl;
import lombok.Data;

/**
 * @author skpeng
 * @version 1.0
 * @description
 * @date 2022/8/14
 */

@Data
public class SBBean1 {

    @ExcelProperty(DealSBExcelServiceImpl.matchColumnName0)
    private String matchColumnName0;
    @ExcelProperty(DealSBExcelServiceImpl.matchColumnName1)
    private String matchColumnName1;
    @ExcelProperty(DealSBExcelServiceImpl.matchColumnName2)
    private String matchColumnName2;
    @ExcelProperty(DealSBExcelServiceImpl.matchColumnName3)
    private String matchColumnName3;
    @ExcelProperty(DealSBExcelServiceImpl.modifyColumnName)
    private String modifyColumnName;

}
