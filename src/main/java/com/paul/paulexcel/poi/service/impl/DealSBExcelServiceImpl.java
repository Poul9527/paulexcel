package com.paul.paulexcel.poi.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.paul.paulexcel.poi.bean.SBBean1;
import com.paul.paulexcel.poi.service.ReadService;
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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author poul9527
 * @version 1.0
 * @description 总数据bigexcel.xlsx，需要修改和筛选excel文件夹D:\excels，里面有多个excel，
 * 部件类型列如果为通用件就用部件类型值和图号1值去总表中匹配，
 * 将长度值赋过来（长度值为相等，则无操作；前提是总表中长度列均有值，不然还要另外处理）
 * 同理专用件与图号2值，沙雕件与图号3值，组合，去总表查，赋值
 * 如果有没能匹配到的则挑出来存入"D:\\bigexcels\\notmatchexcel.xlsx"
 * @date 2022/8/14
 */

@Service
@Slf4j
public class DealSBExcelServiceImpl implements ReadService {
    // 总excel路径
    public static final String bigfilePath = "D:\\bigexcels\\bigexcel.xlsx";
    // 最终没有匹配到的数据汇总excel路径
    public static final String notMatchfilePath = "D:\\bigexcels\\notmatchexcel.xlsx";
    // 要修改文件夹路径
    public static final String folderPath = "D:\\excels";
    // 匹配对应列名
    public static final String matchColumnName0 = "部件类型";
    public static final String matchColumnName1 = "图号1";
    public static final String matchColumnName2 = "图号2";
    public static final String matchColumnName3 = "图号3";
    // 部件类型列的值，分别跟matchColumnName1、2、3列值组合
    public static final String matchColumnValue1 = "通用件";
    public static final String matchColumnValue2 = "专用件";
    public static final String matchColumnValue3 = "沙雕件";
    // 要修改的列名
    public static final String modifyColumnName = "长度";

    public static void main(String[] args) throws IOException {
        new DealSBExcelServiceImpl().readMultiExcel();
    }

    public void readMultiExcel() throws IOException {
        List<SBBean1> bigBean = readBigExcel();
        System.out.println("--------------------读取文件夹，批量解析Excel文件-----------------------");
        File folder = new File(folderPath);

        Map<String, Row> notMatchRowMap = new LinkedHashMap<>();

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

                    List<String> mapKey = new ArrayList<>();

                    int matchColumnNo0 = -1;
                    int matchColumnNo1 = -1;
                    int matchColumnNo2 = -1;
                    int matchColumnNo3 = -1;
                    int modifyColumnNo = -1;
                    //获取表头信息
                    if (firstRow == null) {
                        log.info(String.format("解析{}失败，在第一行表头没有读取到任何数据！", filePath));
                    } else {
                        for (int i1 = 0; i1 < cellCount; i1++) {
                            // 获取对应5个表头所在列号
                            Cell firstRowCell = firstRow.getCell(i1);
                            if (matchColumnName0.equals(getCellVal(firstRowCell))) {
                                matchColumnNo0 = firstRowCell.getColumnIndex();
                            } else if (matchColumnName1.equals(getCellVal(firstRowCell))) {
                                matchColumnNo1 = firstRowCell.getColumnIndex();
                            } else if (matchColumnName2.equals(getCellVal(firstRowCell))) {
                                matchColumnNo2 = firstRowCell.getColumnIndex();
                            } else if (matchColumnName3.equals(getCellVal(firstRowCell))) {
                                matchColumnNo3 = firstRowCell.getColumnIndex();
                            } else if (modifyColumnName.equals(getCellVal(firstRowCell))) {
                                modifyColumnNo = firstRowCell.getColumnIndex();
                            }
                        }
                    }
                    String modifyColumnValue = "";
                    for (int j = 1; j < lastRowCount; j++) {
                        Row row = sheet.getRow(j);
                        String matchColValue0 = getCellVal(row.getCell(matchColumnNo0));
                        if (matchColumnValue1.equals(matchColValue0)) {
                            String columnValue1 = getCellVal(row.getCell(matchColumnNo1));
                            for (SBBean1 sbBean1 : bigBean) {
                                if (matchColValue0.equals(sbBean1.getMatchColumnName0()) && columnValue1.equals(sbBean1.getMatchColumnName1())) {
                                    modifyColumnValue = sbBean1.getModifyColumnName();
                                    break;
                                }
                            }
                        } else if (matchColumnValue2.equals(matchColValue0)) {
                            String columnValue2 = getCellVal(row.getCell(matchColumnNo2));
                            for (SBBean1 sbBean1 : bigBean) {
                                if (matchColValue0.equals(sbBean1.getMatchColumnName0()) && columnValue2.equals(sbBean1.getMatchColumnName2())) {
                                    modifyColumnValue = sbBean1.getModifyColumnName();
                                    break;
                                }
                            }
                        } else if (matchColumnValue3.equals(matchColValue0)) {
                            String columnValue3 = getCellVal(row.getCell(matchColumnNo3));
                            for (SBBean1 sbBean1 : bigBean) {
                                if (matchColValue0.equals(sbBean1.getMatchColumnName0()) && columnValue3.equals(sbBean1.getMatchColumnName3())) {
                                    modifyColumnValue = sbBean1.getModifyColumnName();
                                    break;
                                }
                            }
                        }

                        if (StringUtils.isNotBlank(modifyColumnValue)) {
                            Cell modifyCell = row.getCell(modifyColumnNo);
                            if (modifyCell == null) {
                                row.createCell(modifyColumnNo, CellType.STRING).setCellValue(modifyColumnValue);
                            } else {
                                // 匹配到值一样，不修改
                                if (!modifyColumnValue.equals(getCellVal(modifyCell))) {
                                    row.getCell(modifyColumnNo).setCellValue(modifyColumnValue);
                                }
                            }
                        } else {
                            log.debug("{}行{}数据没匹配到对应值", filePath + "\\" + sheet.getSheetName(), j + 1);
                            // 没匹配到，把这个row的数据写入模板
                            // writeExcel()
                            notMatchRowMap.put(filePath + "\\" + sheet.getSheetName() + "\\" + String.valueOf(j + 1), row);
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

        // 没有匹配到数据需要汇总填入的excel
        File notMatchExcel = new File(notMatchfilePath);
        if (!notMatchExcel.isFile()) {
            notMatchExcel.createNewFile();
        }
        // 无匹配数据写入汇总excel
        Workbook notMatchWB = null;
        FileOutputStream notMatchOS = null;
        try {
            notMatchWB = new XSSFWorkbook();
            Sheet notMatchSh = notMatchWB.createSheet("无匹配数据");
            int number = 1;
            for (Map.Entry<String, Row> entry : notMatchRowMap.entrySet()) {
                String note = entry.getKey();
                Row row = entry.getValue();
                row.createCell(row.getLastCellNum() + 1, CellType.STRING).setCellValue(note);
                Row newRow = notMatchSh.createRow(number);
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    newRow.createCell(j, CellType.STRING).setCellValue(getCellVal(row.getCell(j)));
                }
                number++;
            }
            notMatchOS = new FileOutputStream(notMatchExcel);
            notMatchWB.write(notMatchOS);
        } finally {
            if (notMatchOS != null) {
                notMatchOS.close();
            }
            if (notMatchWB != null) {
                notMatchWB.close();
            }
        }
    }

    /**
     * 获取大Excel的数据
     *
     * @return
     */
    public List<SBBean1> readBigExcel() {
        List<SBBean1> list = new ArrayList<>();
        EasyExcel.read(bigfilePath, SBBean1.class, new ReadListener<SBBean1>() {
            @Override
            public void invoke(SBBean1 o, AnalysisContext analysisContext) {
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
