package com.paul.paulexcel.poi.bean;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.paul.paulexcel.poi.service.impl.DealSBExcelService2Impl;
import lombok.Data;

/**
 * @author skpeng
 * @version 1.0
 * @description     DealSBExcelService2Impl.matchColumnName3中字段为空的需加上@ExcelIgnore
 * @date 2022/8/14
 */

@Data
public class SBBean2 {

//    @ExcelIgnore
    @ExcelProperty(DealSBExcelService2Impl.matchColumnName0)
    private String matchColumnName0;
    @ExcelProperty(DealSBExcelService2Impl.matchColumnName1)
    private String matchColumnName1;
    @ExcelProperty(DealSBExcelService2Impl.matchColumnName2)
    private String matchColumnName2;
    @ExcelIgnore
    @ExcelProperty(DealSBExcelService2Impl.matchColumnName3)
    private String matchColumnName3;
    @ExcelIgnore
    @ExcelProperty(DealSBExcelService2Impl.matchColumnName4)
    private String matchColumnName4;

    @ExcelProperty(DealSBExcelService2Impl.modifyColumnName0)
    private String modifyColumnName0;
    @ExcelIgnore
    @ExcelProperty(DealSBExcelService2Impl.modifyColumnName1)
    private String modifyColumnName1;
    @ExcelIgnore
    @ExcelProperty(DealSBExcelService2Impl.modifyColumnName2)
    private String modifyColumnName2;
    @ExcelIgnore
    @ExcelProperty(DealSBExcelService2Impl.modifyColumnName3)
    private String modifyColumnName3;
    @ExcelIgnore
    @ExcelProperty(DealSBExcelService2Impl.modifyColumnName4)
    private String modifyColumnName4;

}
