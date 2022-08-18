package com.paul.paulexcel.poi.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.paul.paulexcel.poi.bean.SBBean2;
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
 * @description 将modifyColumnName列值为空的行提取出来放入notmatchexcel.xlsx，不支持数据不断往下增加，每执行一次得将该excel移走
 * @date 2022/8/15
 */

@Service
@Slf4j
public class DealSBExcelNoValueServiceImpl implements ReadService {
    // 最终没有匹配到的数据汇总excel路径
    public static final String notMatchfilePath = "D:\\bigexcels\\notmatchexcel.xlsx";
    // 要修改文件夹路径
    public static final String folderPath = "D:\\excels";

    // 修改列名
    public static final String modifyColumnName0 = "MDM编号";
    public static final String modifyColumnName1 = "";
    public static final String modifyColumnName2 = "";
    public static final String modifyColumnName3 = "";
    public static final String modifyColumnName4 = "";

    public static void main(String[] args) throws IOException {
        new DealSBExcelNoValueServiceImpl().wirtenotmatchexcel();
    }

    public static void filesDirs(File file, List<String> filesPath) {
        if (file != null) {
            if (file.isDirectory()) {
                File[] files = file.listFiles();
                for (File flies2 : files) {
                    filesDirs(flies2, filesPath);
                }
            } else if (file.isFile()) {
                String absolutePath = file.getAbsolutePath();
                if (absolutePath.endsWith(".xls") || absolutePath.endsWith(".xlsx")) {
                    filesPath.add(absolutePath);
                }
            }
        } else {
            System.out.println("文件不存在");
        }
    }

    public void wirtenotmatchexcel() throws IOException {
        int noteNo = -1;

        System.out.println("--------------------读取文件夹，批量解析Excel文件-----------------------");
        File folder = new File(folderPath);
        Map<String, Row> notMatchRowMap = new LinkedHashMap<>();

        List<String> filesPath = new ArrayList();
        if (folder.exists()) {
            filesDirs(folder, filesPath);
        } else {
            log.info("文件夹不存在");
            return;
        }
        log.info(String.format("共有excel{%s}个", filesPath.size()));

        // 遍历文件
        for (String filePath : filesPath) {

            Workbook workbook = null;
            FileInputStream fileInputStream = null;
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

                    if(cellCount > noteNo){
                        noteNo = cellCount;
                    }

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
                            if (StringUtils.isBlank(i1CellVal)) {
                                continue;
                            }
                            int i1Index = firstRowCell.getColumnIndex();
                            if (StringUtils.isNotBlank(modifyColumnName0) && modifyColumnName0.equals(i1CellVal)) {
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
                    }

                    for (int j = 1; j <= lastRowCount; j++) {
                        Row row = sheet.getRow(j);

                        if (StringUtils.isNotBlank(modifyColumnName0) && modifyColumnNo0 != -1) {
                            String modifyColumnValue0 = getCellVal(row.getCell(modifyColumnNo0));
                            if (StringUtils.isBlank(modifyColumnValue0)) {
                                notMatchRowMap.put(filePath + "\\" + sheet.getSheetName() + "\\" + String.valueOf(j + 1), row);
                                continue;
                            }
                        }
                        if (StringUtils.isNotBlank(modifyColumnName1) && modifyColumnNo1 != -1) {
                            String modifyColumnValue1 = getCellVal(row.getCell(modifyColumnNo1));
                            if (StringUtils.isBlank(modifyColumnValue1)) {
                                notMatchRowMap.put(filePath + "\\" + sheet.getSheetName() + "\\" + String.valueOf(j + 1), row);
                                continue;
                            }
                        }
                        if (StringUtils.isNotBlank(modifyColumnName2) && modifyColumnNo2 != -1) {
                            String modifyColumnValue2 = getCellVal(row.getCell(modifyColumnNo2));
                            if (StringUtils.isBlank(modifyColumnValue2)) {
                                notMatchRowMap.put(filePath + "\\" + sheet.getSheetName() + "\\" + String.valueOf(j + 1), row);
                                continue;
                            }
                        }
                        if (StringUtils.isNotBlank(modifyColumnName3) && modifyColumnNo3 != -1) {
                            String modifyColumnValue3 = getCellVal(row.getCell(modifyColumnNo3));
                            if (StringUtils.isBlank(modifyColumnValue3)) {
                                notMatchRowMap.put(filePath + "\\" + sheet.getSheetName() + "\\" + String.valueOf(j + 1), row);
                                continue;
                            }
                        }
                        if (StringUtils.isNotBlank(modifyColumnName4) && modifyColumnNo0 != -1) {
                            String modifyColumnValue4 = getCellVal(row.getCell(modifyColumnNo4));
                            if (StringUtils.isBlank(modifyColumnValue4)) {
                                notMatchRowMap.put(filePath + "\\" + sheet.getSheetName() + "\\" + String.valueOf(j + 1), row);
                                continue;
                            }
                        }
                    }

                }

                workbook.close();
                fileInputStream.close();
            } finally {
                try {
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
                row.createCell(noteNo, CellType.STRING).setCellValue(note);
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
