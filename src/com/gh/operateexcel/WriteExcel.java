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
	private static List<String> CELL_HEADS; //��ͷ

    static{
        // ��װ��ʱ������ָ���õ���ͷ��Ϣ��������Ҫ�����Կ������ɶ�̬���ɵ���ͷ
        CELL_HEADS = new ArrayList<>();
        CELL_HEADS.add("ʡ��");
        CELL_HEADS.add("����");
        CELL_HEADS.add("��������");
        CELL_HEADS.add("Ƿ�Ѳ�ѯ");
        CELL_HEADS.add("�µ�");
        CELL_HEADS.add("GOODS_CODE");
        CELL_HEADS.add("֧�����");
        CELL_HEADS.add("֧��������");
        CELL_HEADS.add("���˽��");
    }

    /**
     * ����Excel��д��������Ϣ
     * @param dataList �����б�
     * @return д�����ݺ�Ĺ���������
     * @throws IOException 
     */
//    public static Workbook exportData(List<ExcelDataVO> dataList){
//        // ����xlsx��Excel
//        Workbook workbook = new XSSFWorkbook();
//
//        // ��������xls��Excel����ʹ������Ĺ���������ע��������ʱ�ļ���׺��Ҳ�����Ϊxls
////        Workbook workbook = new HSSFWorkbook();
//
//        // ����Sheet��д���һ�е���ͷ
//        Sheet sheet = buildDataSheet(workbook);
//        //����ÿ�е���������
//        int rowNum = 1;
//        for (Iterator<ExcelDataVO> it = dataList.iterator(); it.hasNext();) {
//            ExcelDataVO data = it.next();
//            if (data == null) {
//                continue;
//            }
//            //���������
//            Row row = sheet.createRow(rowNum++);
//            convertDataToRow(data, row, sheet);
//        }
//        return workbook;
//    }
//    
    public static Workbook exportData(int rowNum,ExcelDataVO data,String patString){
        // ����xlsx��Excel
        Workbook workbook = null;
		try {
			workbook = new XSSFWorkbook();
			if (0==rowNum) {
	        	// ����Sheet��д���һ�е���ͷ
	            Sheet sheet = buildDataSheet(workbook);
	          //���������
	            Row row = sheet.createRow(rowNum+=1);
	            convertDataToRow(data, row, sheet);
			}else {
				FileInputStream fileInputStream = new FileInputStream(patString);
				workbook = new XSSFWorkbook(fileInputStream);
				Sheet sheet = workbook.getSheetAt(0);
				//���������
	            Row row = sheet.createRow(rowNum+=1);
	            convertDataToRow(data, row, sheet);
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        // ��������xls��Excel����ʹ������Ĺ���������ע��������ʱ�ļ���׺��Ҳ�����Ϊxls
//        Workbook workbook = new HSSFWorkbook();

        
        //����ÿ�е���������
//        int rowNum = 1;
        
        
        
        return workbook;
    }
    
    /**
     * ����sheet����д���һ�����ݣ���ͷ��
     * @param workbook ����������
     * @return �Ѿ�д����ͷ��Sheet
     */
    private static Sheet buildDataSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet();
        // ������ͷ���
        for (int i=0; i<CELL_HEADS.size(); i++) {
            sheet.setColumnWidth(i, 4000);
        }
        // ����Ĭ���и�
        sheet.setDefaultRowHeight((short) 400);
        // ����ͷ��Ԫ����ʽ
        CellStyle cellStyle = buildHeadCellStyle(sheet.getWorkbook());
        // д���һ�и��е�����
        Row head = sheet.createRow(0);
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            Cell cell = head.createCell(i);
            cell.setCellValue(CELL_HEADS.get(i));
            cell.setCellStyle(cellStyle);
        }
        return sheet;
    }

    /**
     * ���õ�һ����ͷ����ʽ
     * @param workbook ����������
     * @return ��Ԫ����ʽ����
     */
    private static CellStyle buildHeadCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //���뷽ʽ����
        style.setAlignment(HorizontalAlignment.CENTER);
        //�߿���ɫ�Ϳ������
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // �±߿�
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // ��߿�
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // �ұ߿�
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // �ϱ߿�
        //���ñ�����ɫ
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //����������
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }
    
    /**
     * ����Ƿ�Ѳ�ѯʧ�ܣ��µ�ʧ�ܣ�goods_code������ʱ����ʽ
     */
    
    private static CellStyle buildCellStyle(Workbook workbook) {
    	CellStyle cellStyle = workbook.createCellStyle();
    	// ���ö��뷽ʽ
    	cellStyle.setAlignment(HorizontalAlignment.LEFT);
    	// ���ñ߿���ɫ�Ϳ��
    	cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // �±߿�
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // ��߿�
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex()); // �ұ߿�
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex()); // �ϱ߿�
        // ���ñ�����ɫ
        cellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        return cellStyle;
    }
    
    /**
     * ������ת������
     * @param data Դ����
     * @param row �ж���
     * @return
     */
    private static void convertDataToRow(ExcelDataVO data, Row row,Sheet sheet){
        int cellNum = 0;
        Cell cell;        
        // ʡ��
        cell = row.createCell(cellNum++);
        cell.setCellValue(null == data.getProvince() ? "" : data.getProvince());
        // ����
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getConsNo());
        // ��������
        cell = row.createCell(cellNum++);
        if (!("��ѹ".equals(data.getConsSortCode()) || "��ѹ�Ǿ�".equals(data.getConsSortCode()) || "��ѹ����".equals(data.getConsSortCode()))) {
			cell.setCellValue("");
		}else {
	        cell.setCellValue((null == data.getConsSortCode() || "".equals(data.getConsSortCode())) ? "" : data.getConsSortCode());
		}
        // Ƿ�Ѳ�ѯ
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getSearch());
        // �µ�
        cell = row.createCell(cellNum++);  
        cell.setCellValue(data.getSaleOrder());
        setRowCellStyle(data.getSaleOrder(), row, sheet);
        
        // goods_code
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getGoodsCode());
        // ֧�����
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getPay());
        setRowCellStyle(data.getPay(), row, sheet);
        // ֧��������
        cell = row.createCell(cellNum++);
        cell.setCellValue(data.getOrderNo());
        // ���˽��
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
