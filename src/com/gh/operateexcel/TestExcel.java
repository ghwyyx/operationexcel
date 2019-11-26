package com.gh.operateexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.gh.operateexcel.WriteExcel;


public class TestExcel {
	
	private static Logger logger = Logger.getLogger(TestExcel.class.getName());

    public static void main(String[] args) {
//        // ������Ҫд��������б�
//        List<ExcelDataVO> dataVOList = new ArrayList<>();
//        ExcelDataVO dataVO = new ExcelDataVO();
//        dataVO.setProvince("�ӱ�");
//        dataVO.setConsNo("1234567890");
//        dataVO.setConsSortCode("");
//        dataVO.setSearch("���׳ɹ�");
//        dataVO.setSaleOrder("���׳ɹ�");
//        dataVO.setGoodsCode("C0010001");
//        dataVO.setPay("���׳ɹ�");
//        dataVO.setOrderNo("320191119110095344449");
//        dataVO.setYxBackResult("1");
//        ExcelDataVO dataVO2 = new ExcelDataVO();
//        dataVO2.setProvince("�ӱ�");
//        dataVO2.setConsNo("1234567890");
//        dataVO2.setConsSortCode("01");
//        dataVO2.setSearch("���׳ɹ�");
//        dataVO2.setSaleOrder("");
//        dataVO2.setGoodsCode("C0010001");
//        dataVO2.setPay("���׳ɹ�");
//        dataVO2.setOrderNo("320191119110095344449");
//        dataVO2.setYxBackResult("1");
//        dataVOList.add(dataVO);
//        dataVOList.add(dataVO2);
//
//        // д�����ݵ�������������
//        Workbook workbook = WriteExcel.exportData(dataVOList);
//
//        // ���ļ�����ʽ�������������
//        FileOutputStream fileOut = null;
//        try {
//            String exportFilePath = "C:\\Users\\guhao\\Desktop\\27\\result\\writeExample.xlsx";
//            File exportFile = new File(exportFilePath);
//            if (!exportFile.exists()) {
//                exportFile.createNewFile();
//            }
//
//            fileOut = new FileOutputStream(exportFilePath);
//            workbook.write(fileOut);
//            fileOut.flush();
//        } catch (Exception e) {
//            logger.warning("���Excelʱ�������󣬴���ԭ��" + e.getMessage());
//        } finally {
//            try {
//                if (null != fileOut) {
//                    fileOut.close();
//                }
//                if (null != workbook) {
//                    workbook.close();
//                }
//            } catch (IOException e) {
//                logger.warning("�ر������ʱ�������󣬴���ԭ��" + e.getMessage());
//            }
//        }
//        logger.info("ִ�����");
        
    	for (int i = 0; i < 2; i++) {
    		// ������Ҫд��������б�
            ExcelDataVO dataVO = new ExcelDataVO();
            dataVO.setProvince("�ӱ�");
            dataVO.setConsNo("1234567890");
            dataVO.setConsSortCode("");
            dataVO.setSearch("���׳ɹ�");
            dataVO.setSaleOrder("���׳ɹ�");
            dataVO.setGoodsCode("C0010001");
            dataVO.setPay("���׳ɹ�");
            dataVO.setOrderNo("320191119110095344449");
            dataVO.setYxBackResult("1");
            
            String exportFilePath = "C:\\Users\\guhao\\Desktop\\27\\result\\writeExample.xlsx";
            
         // д�����ݵ�������������
            Workbook workbook = WriteExcel.exportData(i,dataVO,exportFilePath);
            
         // ���ļ�����ʽ�������������
            FileOutputStream fileOut = null;
            try {
                
                File exportFile = new File(exportFilePath);
                if (!exportFile.exists()) {
                    exportFile.createNewFile();
                }

                fileOut = new FileOutputStream(exportFilePath);
                workbook.write(fileOut);
                fileOut.flush();
            } catch (Exception e) {
                logger.warning("���Excelʱ�������󣬴���ԭ��" + e.getMessage());
            } finally {
                try {
                    if (null != fileOut) {
                        fileOut.close();
                    }
                    if (null != workbook) {
                        workbook.close();
                    }
                } catch (IOException e) {
                    logger.warning("�ر������ʱ�������󣬴���ԭ��" + e.getMessage());
                }
            }
            logger.info("ִ�����");
            
		}
    	 
       

        

        
        
        
        
    }
    
    public static boolean fileExist(String filePath) {
		boolean flag = false;
		File file = new File(filePath);
		flag = file.exists();
		return flag;
	}
}
