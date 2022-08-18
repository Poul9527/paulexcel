package com.paul.paulexcel.poi.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.paul.paulexcel.poi.bean.SBBean2;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author poul9527
 * @version 1.0
 * @description 多个匹配列，修改多个列，上限5
 * @date 2022/8/15
 */

@Service
@Slf4j
public class DealSBExcelService2Impl {
    // 匹配总excel路径
    public static final String bigfilePath = "D:\\bigexcels\\jgj.xlsx";
    // 要修改文件夹路径
    public static final String folderPath = "D:\\excels";
    // 用来匹配列名
    public static final String matchColumnName0 = "代号";
    public static final String matchColumnName1 = "名称";
    public static final String matchColumnName2 = "材料";
    public static final String matchColumnName3 = "";
    public static final String matchColumnName4 = "";
    // 修改列名
    public static final String modifyColumnName0 = "MDM编号";
    public static final String modifyColumnName1 = "";
    public static final String modifyColumnName2 = "";
    public static final String modifyColumnName3 = "";
    public static final String modifyColumnName4 = "";

    public static void main(String[] args) throws IOException {
        new DealSBExcelService2Impl().readMultiExcel();
    }

    public void readMultiExcel() throws IOException {
        List<SBBean2> bigBean = readBigExcel();
        System.out.println("--------------------读取文件夹，批量解析Excel文件-----------------------");
        File folder = new File(folderPath);

        List<String> filesPath = new ArrayList();
        if (folder.exists()) {
            File[] files = folder.listFiles();
            for (File file2 : files) {
                String absolutePath = file2.getAbsolutePath();
                if (file2.isFile() && (absolutePath.endsWith(".xls") || absolutePath.endsWith(".xlsx"))) {
                    filesPath.add(absolutePath);
                }
            }
        } else {
            log.info("文件夹不存在");
            return;
        }
        log.info(String.format("共有excel{%s}个", filesPath.size()));

        // 遍历文件
        for (String filePath : filesPath) {

            Workbook workbook = null;
            FileInputStream fileInputStream = null;
            FileOutputStream out = null;
            try {
                File excelFile = new File(filePath);
                fileInputStream = new FileInputStream(excelFile);
                if (filePath.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fileInputStream);
                } else if (filePath.endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fileInputStream);
                }
                // 获取一个Excel中sheet数量
                int sheetCount = workbook.getNumberOfSheets();
                for (int i = 0; i < sheetCount; i++) {
                    Sheet sheet = workbook.getSheetAt(i);

                    if (sheet == null) {
                        continue;
                    }
                    //获取第一行的序号
                    int firstRowCount = sheet.getFirstRowNum();
                    Row firstRow = sheet.getRow(firstRowCount);
                    //获取列数
                    int cellCount = firstRow.getLastCellNum();
                    int lastRowCount = sheet.getLastRowNum();

                    // 匹配列号
                    int matchColumnNo0 = -1;
                    int matchColumnNo1 = -1;
                    int matchColumnNo2 = -1;
                    int matchColumnNo3 = -1;
                    int matchColumnNo4 = -1;

                    //要修改填入的列号
                    int modifyColumnNo0 = -1;
                    int modifyColumnNo1 = -1;
                    int modifyColumnNo2 = -1;
                    int modifyColumnNo3 = -1;
                    int modifyColumnNo4 = -1;

                    //获取表头信息
                    if (firstRow == null) {
                        log.info(String.format("解析{}失败，在第一行表头没有读取到任何数据！", filePath));
                    } else {
                        for (int i1 = 0; i1 < cellCount; i1++) {
                            // 获取对应表头所在列号
                            Cell firstRowCell = firstRow.getCell(i1);
                            String i1CellVal = getCellVal(firstRowCell);
                            int i1Index = i1;
                            if (StringUtils.isNotBlank(matchColumnName0) && matchColumnName0.equals(i1CellVal)) {
                                matchColumnNo0 = i1Index;
                            } else if (StringUtils.isNotBlank(matchColumnName1) && matchColumnName1.equals(i1CellVal)) {
                                matchColumnNo1 = i1Index;
                            } else if (StringUtils.isNotBlank(matchColumnName2) && matchColumnName2.equals(i1CellVal)) {
                                matchColumnNo2 = i1Index;
                            } else if (StringUtils.isNotBlank(matchColumnName3) && matchColumnName3.equals(i1CellVal)) {
                                matchColumnNo3 = i1Index;
                            } else if (StringUtils.isNotBlank(matchColumnName4) && matchColumnName4.equals(i1CellVal)) {
                                matchColumnNo4 = i1Index;
                            } else if (StringUtils.isNotBlank(modifyColumnName0) && modifyColumnName0.equals(i1CellVal)) {
                                modifyColumnNo0 = i1Index;
                            } else if (StringUtils.isNotBlank(modifyColumnName1) && modifyColumnName1.equals(i1CellVal)) {
                                modifyColumnNo1 = i1Index;
                            } else if (StringUtils.isNotBlank(modifyColumnName2) && modifyColumnName2.equals(i1CellVal)) {
                                modifyColumnNo2 = i1Index;
                            } else if (StringUtils.isNotBlank(modifyColumnName3) && modifyColumnName3.equals(i1CellVal)) {
                                modifyColumnNo3 = i1Index;
                            } else if (StringUtils.isNotBlank(modifyColumnName4) && modifyColumnName4.equals(i1CellVal)) {
                                modifyColumnNo4 = i1Index;
                            }
                        }
                        if (StringUtils.isNotBlank(modifyColumnName0) && modifyColumnNo0 == -1) {
                            firstRow.createCell(firstRow.getLastCellNum() + 1).setCellValue(modifyColumnName0);
                            modifyColumnNo0 = firstRow.getLastCellNum() + 1;
                        }
                        if (StringUtils.isNotBlank(modifyColumnName1) && modifyColumnNo1 == -1) {
                            firstRow.createCell(firstRow.getLastCellNum() + 1).setCellValue(modifyColumnName1);
                            modifyColumnNo1 = firstRow.getLastCellNum() + 1;

                        }
                        if (StringUtils.isNotBlank(modifyColumnName2) && modifyColumnNo2 == -1) {
                            firstRow.createCell(firstRow.getLastCellNum() + 1).setCellValue(modifyColumnName2);
                            modifyColumnNo2 = firstRow.getLastCellNum() + 1;
                        }
                        if (StringUtils.isNotBlank(modifyColumnName3) && modifyColumnNo3 == -1) {
                            firstRow.createCell(firstRow.getLastCellNum() + 1).setCellValue(modifyColumnName3);
                            modifyColumnNo3 = firstRow.getLastCellNum() + 1;
                        }
                        if (StringUtils.isNotBlank(modifyColumnName4) && modifyColumnNo4 == -1) {
                            firstRow.createCell(firstRow.getLastCellNum() + 1).setCellValue(modifyColumnName4);
                            modifyColumnNo4 = firstRow.getLastCellNum() + 1;
                        }
                    }

                    for (int j = 1; j <= lastRowCount; j++) {
                        Row row = sheet.getRow(j);
                        boolean isMatchValue = false;
                        SBBean2 sbBean = new SBBean2();
                        for (SBBean2 sbBean2 : bigBean) {
                            isMatchValue = false;
                            if (matchColumnNo0 != -1) {
                                String matchColValue0 = getCellVal(row.getCell(matchColumnNo0));
                                String bigMatchColValue0 = sbBean2.getMatchColumnName0();
                                if (matchColValue0.equals(bigMatchColValue0)) {
                                    log.info(sbBean2.getMatchColumnName0() + sbBean2.getMatchColumnName1() + sbBean2.getMatchColumnName2());
                                    isMatchValue = true;
                                } else {
                                    isMatchValue = false;
                                    continue;
                                }
                            }
                            if (matchColumnNo1 != -1) {
                                String matchColValue1 = getCellVal(row.getCell(matchColumnNo1));
                                String bigMatchColValue1 = sbBean2.getMatchColumnName1();
                                if (matchColValue1.equals(bigMatchColValue1)) {
                                    isMatchValue = true;
                                } else {
                                    isMatchValue = false;
                                    continue;
                                }
                            }
                            if (matchColumnNo2 != -1) {
                                String matchColValue2 = getCellVal(row.getCell(matchColumnNo2));
                                String bigMatchColValue2 = sbBean2.getMatchColumnName2();
                                if (matchColValue2.equals(bigMatchColValue2)) {
                                    isMatchValue = true;
                                } else {
                                    isMatchValue = false;
                                    continue;
                                }
                            }
                            if (matchColumnNo3 != -1) {
                                String matchColValue3 = getCellVal(row.getCell(matchColumnNo3));
                                String bigMatchColValue3 = sbBean2.getMatchColumnName3();
                                if (matchColValue3.equals(bigMatchColValue3)) {
                                    isMatchValue = true;
                                } else {
                                    isMatchValue = false;
                                    continue;
                                }
                            }
                            if (matchColumnNo4 != -1) {
                                String matchColValue4 = getCellVal(row.getCell(matchColumnNo4));
                                String bigMatchColValue4 = sbBean2.getMatchColumnName4();
                                if (matchColValue4.equals(bigMatchColValue4)) {
                                    isMatchValue = true;
                                } else {
                                    isMatchValue = false;
                                    continue;
                                }
                            }
                            if (isMatchValue) {
                                sbBean = sbBean2;
                                break;
                            }
                        }

                        if (isMatchValue) {
                            if (modifyColumnNo0 != -1) {
                                String bigModifyColumnValue0 = sbBean.getModifyColumnName0();
                                if (StringUtils.isNotBlank(bigModifyColumnValue0)) {
                                    Cell modifyCell = row.getCell(modifyColumnNo0);
                                    if (modifyCell == null) {
                                        row.createCell(modifyColumnNo0, CellType.STRING).setCellValue(bigModifyColumnValue0);
                                    } else {
                                        // 匹配到值一样，不修改
                                        if (!getCellVal(modifyCell).equals(bigModifyColumnValue0)) {
                                            row.getCell(modifyColumnNo0).setCellValue(bigModifyColumnValue0);
                                        }
                                    }
                                }
                            }
                            if (modifyColumnNo1 != -1) {
                                String bigModifyColumnValue1 = sbBean.getModifyColumnName1();
                                if (StringUtils.isNotBlank(bigModifyColumnValue1)) {
                                    Cell modifyCell = row.getCell(modifyColumnNo1);
                                    if (modifyCell == null) {
                                        row.createCell(modifyColumnNo1, CellType.STRING).setCellValue(bigModifyColumnValue1);
                                    } else {
                                        // 匹配到值一样，不修改
                                        if (!getCellVal(modifyCell).equals(bigModifyColumnValue1)) {
                                            row.getCell(modifyColumnNo1).setCellValue(bigModifyColumnValue1);
                                        }
                                    }
                                }
                            }
                            if (modifyColumnNo2 != -1) {
                                String bigModifyColumnValue2 = sbBean.getModifyColumnName2();
                                if (StringUtils.isNotBlank(bigModifyColumnValue2)) {
                                    Cell modifyCell = row.getCell(modifyColumnNo2);
                                    if (modifyCell == null) {
                                        row.createCell(modifyColumnNo2, CellType.STRING).setCellValue(bigModifyColumnValue2);
                                    } else {
                                        // 匹配到值一样，不修改
                                        if (!getCellVal(modifyCell).equals(bigModifyColumnValue2)) {
                                            row.getCell(modifyColumnNo2).setCellValue(bigModifyColumnValue2);
                                        }
                                    }
                                }
                            }
                            if (modifyColumnNo3 != -1) {
                                String bigModifyColumnValue3 = sbBean.getModifyColumnName3();
                                if (StringUtils.isNotBlank(bigModifyColumnValue3)) {
                                    Cell modifyCell = row.getCell(modifyColumnNo3);
                                    if (modifyCell == null) {
                                        row.createCell(modifyColumnNo3, CellType.STRING).setCellValue(bigModifyColumnValue3);
                                    } else {
                                        // 匹配到值一样，不修改
                                        if (!getCellVal(modifyCell).equals(bigModifyColumnValue3)) {
                                            row.getCell(modifyColumnNo3).setCellValue(bigModifyColumnValue3);
                                        }
                                    }
                                }
                            }
                            if (modifyColumnNo4 != -1) {
                                String bigModifyColumnValue4 = sbBean.getModifyColumnName4();
                                if (StringUtils.isNotBlank(bigModifyColumnValue4)) {
                                    Cell modifyCell = row.getCell(modifyColumnNo4);
                                    if (modifyCell == null) {
                                        row.createCell(modifyColumnNo4, CellType.STRING).setCellValue(bigModifyColumnValue4);
                                    } else {
                                        // 匹配到值一样，不修改
                                        if (!getCellVal(modifyCell).equals(bigModifyColumnValue4)) {
                                            row.getCell(modifyColumnNo4).setCellValue(bigModifyColumnValue4);
                                        }
                                    }
                                }
                            }
                        } else {
                            log.debug("{}行{}数据没匹配到", filePath + "\\" + sheet.getSheetName(), j + 1);
                        }

                    }

                }

                // 写入excel
                out = new FileOutputStream(filePath);
                workbook.write(out);

                out.close();
                workbook.close();
                fileInputStream.close();
            } finally {
                try {
                    if (null != out) {
                        out.close();
                    }
                    if (null != workbook) {
                        workbook.close();
                    }
                    if (null != fileInputStream) {
                        fileInputStream.close();
                    }
                } catch (Exception e) {
                    log.error("关闭数据流出错，" + filePath + "前面数据已被修改，请重新使用原来文件！错误信息：", e);
                    return;
                }
            }

        }
    }

    /**
     * 获取大Excel的数据
     *
     * @return
     */
    public List<SBBean2> readBigExcel() {
        List<SBBean2> list = new ArrayList<>();
        EasyExcel.read(bigfilePath, SBBean2.class, new ReadListener<SBBean2>() {
            @Override
            public void invoke(SBBean2 o, AnalysisContext analysisContext) {
                list.add(o);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext analysisContext) {

            }
        }).sheet().doRead();

        return list;
    }

    /**
     * 获取单元格的值
     *
     * @param cel
     * @return
     */
    public static String getCellVal(Cell cel) {
        if (cel == null) {
            return "";
        }
        if (cel.getCellType() == Cell.CELL_TYPE_STRING) {
            return cel.getRichStringCellValue().getString();
        }
        if (cel.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return cel.getNumericCellValue() + "";
        }
        if (cel.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return cel.getBooleanCellValue() + "";
        }
        if (cel.getCellType() == Cell.CELL_TYPE_FORMULA) {
            return cel.getCellFormula() + "";
        }
        return cel.toString();
    }

}
