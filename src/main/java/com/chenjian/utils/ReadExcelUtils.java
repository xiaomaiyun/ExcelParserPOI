package com.chenjian.utils;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

public class ReadExcelUtils {
    public static String readExcelFilePath(String filePath) {
        String content = "";
        try {
            if (!WDWUtil.validateExcel(filePath)) {
                content = "文件名：" + filePath + "，Excel文件格式错误，请更正后再尝试!";

            } else {
                if (WDWUtil.isExcel2003(filePath)) {
                    content=readExcel2003(filePath);

                } else if (WDWUtil.isExcel2007(filePath)) {
                    content=readExcel2007(filePath);

                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return content;
    }

    public static String readExcel2003(String filePath) throws IOException {
        StringBuilder content = new StringBuilder();
        // 创建对Excel工作簿文件的引用
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(filePath));

        //获取一个单元格对象，用于保存每个单元格的信息
        ContentModle contentModle = new ContentModle();
        //保存文件路径
        contentModle.setFilePath(filePath);

        System.out.println("Sheet总数量：" + workbook.getNumberOfSheets());
        //遍历每个sheet
        for (int numSheets = 0; numSheets < workbook.getNumberOfSheets(); numSheets++) {
            //保存Sheet的名字
            contentModle.setSheetName(workbook.getSheetName(numSheets));

            if (null != workbook.getSheetAt(numSheets)) {
                // 获得一个sheet
                HSSFSheet aSheet = workbook.getSheetAt(numSheets);
                //行的数量必须包含等于，否则最后一行没有读取，因为getLastRowNum()方法表示获取表单中最后一行的索引（而不是总行数）
                System.out.println(contentModle.getSheetName()+"总行数：" + aSheet.getLastRowNum()+1);
                for (int rowNumOfSheet = 0; rowNumOfSheet <= aSheet.getLastRowNum(); rowNumOfSheet++) {
                    // 遍历每一行
                    if (null != aSheet.getRow(rowNumOfSheet)) {
                        // 获得一行
                        HSSFRow aRow = aSheet.getRow(rowNumOfSheet);

                        //getLastCellNum()获取此行中包含的最后一个单元格的总列数(而不是索引)
                        System.out.println("总列数：" + aRow.getLastCellNum());
                        for (int cellNumOfRow = 0; cellNumOfRow < aRow.getLastCellNum(); cellNumOfRow++) {
                            //遍历每个单元格
                            if (null != aRow.getCell(cellNumOfRow)) {

                                //获取单元格的值
                                HSSFCell aCell = aRow.getCell(cellNumOfRow);

                                //单元格行索引+1
                                contentModle.setRowNum(aCell.getRowIndex() + 1);
                                //单元格列索引+1
                                contentModle.setCellNumOfRow(aCell.getColumnIndex() + 1);
                                //单元格内容
                                contentModle.setContent(convert(aCell));

                            } else {
                                contentModle.setRowNum(rowNumOfSheet + 1);
                                contentModle.setCellNumOfRow(cellNumOfRow + 1);
                                contentModle.setContent(null);
                            }

//                            System.out.print("["+""+contentModle.getFilePath()+","+contentModle.getSheetName()+",("+contentModle.getRowNum()+","+contentModle.getCellNumOfRow()+")]"+contentModle.getContent());
                            System.out.print("(" + contentModle.getRowNum() + "," + contentModle.getCellNumOfRow() + ")" + contentModle.getContent());
                            System.out.print(" ");
                            content.append(contentModle.getContent());
                            content.append("\t");
                        }
                        System.out.println();
                        content.append("\n");
                    }
                }
            }
        }
        return content.toString();

    }

    public static String readExcel2007(String filePath) throws IOException {
        StringBuilder content = new StringBuilder();
        // 创建对Excel工作簿文件的引用

        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(filePath));

        //获取一个单元格对象，用于保存每个单元格的信息
        ContentModle contentModle = new ContentModle();
        //保存文件路径
        contentModle.setFilePath(filePath);

        System.out.println("Sheet的数量：" + workbook.getNumberOfSheets());
        //遍历每个sheet
        for (int numSheets = 0; numSheets < workbook.getNumberOfSheets(); numSheets++) {
            //保存Sheet的名字
            contentModle.setSheetName(workbook.getSheetName(numSheets));

            if (null != workbook.getSheetAt(numSheets)) {
                // 获得一个sheet
                XSSFSheet aSheet = workbook.getSheetAt(numSheets);
                //行的数量必须包含等于，否则最后一行没有读取，因为getLastRowNum()方法表示获取表单中最后一行的索引（而不是总行数）
                System.out.println("LastRowNum的值：" + aSheet.getLastRowNum());
                for (int rowNumOfSheet = 0; rowNumOfSheet <= aSheet.getLastRowNum(); rowNumOfSheet++) {
                    // 遍历每一行
                    if (null != aSheet.getRow(rowNumOfSheet)) {
                        // 获得一行
                        XSSFRow aRow = aSheet.getRow(rowNumOfSheet);

                        //getLastCellNum()获取此行中包含的最后一个单元格的总列数(而不是索引)
                        System.out.println("LastCellNum的值：" + aRow.getLastCellNum());
                        for (int cellNumOfRow = 0; cellNumOfRow < aRow.getLastCellNum(); cellNumOfRow++) {
                            //遍历每个单元格
                            if (null != aRow.getCell(cellNumOfRow)) {

                                //获取单元格的值
                                XSSFCell aCell = aRow.getCell(cellNumOfRow);

                                //单元格行索引+1
                                contentModle.setRowNum(aCell.getRowIndex() + 1);
                                //单元格列索引+1
                                contentModle.setCellNumOfRow(aCell.getColumnIndex() + 1);
                                //单元格内容
                                contentModle.setContent(convert(aCell));

                            } else {
                                contentModle.setRowNum(rowNumOfSheet + 1);
                                contentModle.setCellNumOfRow(cellNumOfRow + 1);
                                contentModle.setContent(null);
                            }

//                            System.out.print("["+""+contentModle.getFilePath()+","+contentModle.getSheetName()+",("+contentModle.getRowNum()+","+contentModle.getCellNumOfRow()+")]"+contentModle.getContent());
                            System.out.print("(" + contentModle.getRowNum() + "," + contentModle.getCellNumOfRow() + ")" + contentModle.getContent());
                            System.out.print(" ");
                            content.append(contentModle.getContent());
                            content.append("\t");
                        }
                        System.out.println();
                        content.append("\n");
                    }
                }
            }
        }
        return content.toString();

    }


    private static String convert(Cell cell) {
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

        String cellValue = null;
        if (cell == null) {
            return cellValue;
        }
        switch (cell.getCellType()) {
            //文本
            case Cell.CELL_TYPE_STRING:
                cellValue = cell.getStringCellValue();
                break;
            //布尔类型
            case Cell.CELL_TYPE_BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            //数字，日期
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    //日期类型
                    cellValue = format.format(cell.getDateCellValue());
                } else {
                    //数字类型
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                break;
            //空白
            case Cell.CELL_TYPE_BLANK:
                cellValue = cell.getStringCellValue();
                break;
            //公式
            case Cell.CELL_TYPE_FORMULA:
                cellValue = cell.getCellFormula();
                break;
            //错误
            case Cell.CELL_TYPE_ERROR:
//                cellValue = String.valueOf(cell.getErrorCellValue());
                cellValue = null;
                break;
            default:
                cellValue = null;

        }
        return cellValue;
    }

}
