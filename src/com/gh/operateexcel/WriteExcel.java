package com.gh.operateexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.SheetBuilder;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;



public class WriteExcel {
	private static List<String> CELL_HEADS; //列头

    static{
        // 类装载时就载入指定好的列头信息，如有需要，可以考虑做成动态生成的列头
        CELL_HEADS = new ArrayList<>();
        CELL_HEADS.add("省码");
        CELL_HEADS.add("户号");
        CELL_HEADS.add("户号类型");
        CELL_HEADS.add("欠费查询");
        CELL_HEADS.add("下单");
        CELL_HEADS.add("GOODS_CODE");
        CELL_HEADS.add("支付结果");
        CELL_HEADS.add("支付订单号");
        CELL_HEADS.add("销账结果");
    }

    /**
     * 生成Excel并写入数据信息
     * @param dataList 数据列表
     * @return 写入数据后的工作簿对象
     * @throws IOException 
     */
//    public static Workbook exportData(List<ExcelDataVO> dataList){
//        // 生成xlsx的Excel
//        Workbook workbook = new XSSFWorkbook();
//
//        // 如需生成xls的Excel，请使用下面的工作簿对象，注意后续输出时文件后缀名也需更改为xls
////        Workbook workbook = new HSSFWorkbook();
//
//        // 生成Sheet表，写入第一行的列头
//        Sheet sheet = buildDataSheet(workbook);
//        //构建每行的数据内容
//        int rowNum = 1;
//        for (Iterator<ExcelDataVO> it = dataList.iterator(); it.hasNext();) {
//            ExcelDataVO data = it.next();
//            if (data == null) {
//                continue;
//            }
//            //输出行数据
//            Row row = sheet.createRow(rowNum++);
//            convertDataToRow(data, row, sheet);
//        }
//        return workbook;
//    }
//    
    public static Workbook exportData(int rowNum,ExcelDataVO data,String patString){
        // 生成xlsx的Excel
        Workbook workbook = null;
		try {
			workbook = new XSSFWorkbook();
			if (0==rowNum) {
	        	// 生成Sheet表，写入第一行的列头
	            Sheet sheet = buildDataSheet(workbook);
	          //输出行数据
	            Row row = sheet.createRow(rowNum+=1);
	            convertDataToRow(data, row, sheet);
			}else {
				FileInputStream fileInputStream = new FileInputStream(patString);
				workbook = new XSSFWorkbook(fileInputStream);
				Sheet sheet = workbook.getSheetAt(0);
				//输出行数据
	            Row row = sheet.createRow(rowNum+=1);
	            convertDataToRow(data, row, sheet);
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        // 如需生成xls的Excel，请使用下面的工作簿对象，注意后续输出时文件后缀名也需更改为xls
//        Workbook workbook = new HSSFWorkbook();

        
        //构建每行的数据内容
//        int rowNum = 1;
        
        
        
        return workbook;
    }
    
    /**
     * 生成sheet表，并写入第一行数据（列头）
     * @param workbook 工作簿对象
     * @return 已经写入列头的Sheet
     */
    private static Sheet buildDataSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet();
        // 设置列头宽度
        for (int i=0; i<CELL_HEADS.size(); i++) {
            sheet.setColumnWidth(i, 4000);
        }
        // 设置默认行高
        sheet.setDefaultRowHeight((short) 400);
        // 构建头单元格样式
        CellStyle cellStyle = buildHeadCellStyle(sheet.getWorkbook());
        // 写入第一行各列的数据
        Row head = sheet.createRow(0);
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            Cell cell = head.createCell(i);
            cell.setCellValue(CELL_HEADS.get(i));
            cell.setCellStyle(cellStyle);
        }
        return sheet;
    }

    /**
     * 设置第一行列头的样式
     * @param workbook 工作簿对象
     * @return 单元格样式对象
     */
    private static CellStyle buildHeadCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //对齐方式设置
        style.setAlignment(HorizontalAlignment.CENTER);
        //边框颜色和宽度设置
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下边框
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框
        //设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //粗体字设置
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }
    
    /**
     * 设置欠费查询失败，下单失败，goods_code不符合时的样式
     */
    
    private static CellStyle buildCellStyle(Workbook workbook) {
    	CellStyle cellStyle = workbook.createCellStyle();
    	// 设置对齐方式
    	cellStyle.setAlignment(HorizontalAlignment.LEFT);
    	// 设置边框颜色和宽度
    	cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框
        // 设置背景颜色
        cellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        return cellStyle;
    }
    
    /**
     * 将数据转换成行
     * @param data 源数据
     * @param row 行对象
     * @return
     */
    private static void convertDataToRow(ExcelDataVO data, Row row,Sheet sheet){
        int cellNum = 0;
        Cell cell;        
        // 省码
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getProvince() ? "" : data.getProvince());
        // 户号
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getConsNo());
        // 户号类型
        cell = row.createCell(cellNum++);
        if (!("高压".equals(data.getConsSortCode()) || "低压非居".equals(data.getConsSortCode()) || "低压居民".equals(data.getConsSortCode()))) {
			cell.setCellValue("");
		}else {
	        cell.setCellValue((null == data.getConsSortCode() || "".equals(data.getConsSortCode())) ? "" : data.getConsSortCode());
		}
        // 欠费查询
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getSearch());
        // 下单
        cell = row.createCell(cellNum++);  
        cell.setCellValue(data.getSaleOrder());
        setRowCellStyle(data.getSaleOrder(), row, sheet);
        
        // goods_code
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getGoodsCode());
        // 支付结果
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getPay());
        setRowCellStyle(data.getPay(), row, sheet);
        // 支付订单号
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getOrderNo());
        // 销账结果
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getYxBackResult());
    }
    
    private static void setRowCellStyle(String string,Row row,Sheet sheet) {
    	if("".equals(string) || null == string) {
    		for (Cell cell2 : row) {
    			cell2.setCellStyle(buildCellStyle(sheet.getWorkbook()));
    		}
        }
    	
	}
}
