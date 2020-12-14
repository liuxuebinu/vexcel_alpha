package org.vexcel.tools;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.vexcel.engine.RuleEngine;
import org.vexcel.exception.ValidateRuntimeException;
import org.vexcel.exception.ValidateXmlException;
import org.vexcel.pojo.*;

public class ExcelUtils {


    private static String getCellText(org.apache.poi.ss.usermodel.Cell cell) {
        String celltext = "";
        if(cell != null) {
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            celltext = cell.getStringCellValue();
            if(celltext == null){
                celltext = "";
            }
        }
        return celltext;
    }


    private static boolean  checkHSSFRowIsEmpty(HSSFRow hssfRow){
        List<String> rowCells = new ArrayList<>();
        java.util.Iterator<org.apache.poi.ss.usermodel.Cell> cellIt = hssfRow.cellIterator();
        while(cellIt.hasNext()){
            Cell cell = cellIt.next();
            if(cell !=null){
                String cellText = getCellText(cell);
                if(!CommonUtil.isNull(cellText)){
                    rowCells.add(cellText);
                    return false;
                }
            }
        }
        return (rowCells.size()<=0);
    }

    private static boolean  checkXSSFRowIsEmpty(XSSFRow xssfRow){

        List<String> rowCells = new ArrayList<>();
        java.util.Iterator<org.apache.poi.ss.usermodel.Cell> cellIt = xssfRow.cellIterator();
        while(cellIt.hasNext()){
            Cell cell = cellIt.next();
            if(cell !=null){
                String cellText = getXssCellText(cell);
                if(!CommonUtil.isNull(cellText)){
                    rowCells.add(cellText);
                    return false;
                }
            }
        }
        return (rowCells.size()<=0);
    }

    private static String getXssCellText(org.apache.poi.ss.usermodel.Cell cell) {
        String celltext = "";
        if(cell != null) {
            cell.setCellType(XSSFCell.CELL_TYPE_STRING);
            celltext = cell.getStringCellValue();
            if(celltext == null){
                celltext = "";
            }
        }
        return celltext;

    }

    public static ValidateResult readExcel(InputStream is,List<VSheet> rules,
                                           String excelType) {
        if ("xls".equals(excelType)) {
            return readExcel_XLS(is, rules,excelType);
        } else {
            return readExcel_XLSX(is, rules,excelType);
        }

    }



    public static ValidateResult readExcel_XLS(InputStream is, List<VSheet> rules,
            String excelType) {
        Integer excelCounts = 0;
        int count = 0;
        ValidateResult result = new ValidateResult();
        result.setSuccess(true);
        StringBuilder msgs = new StringBuilder();
        result.setErrorMsg(msgs);

        HSSFWorkbook hssfworkbook = null;
        try {
            hssfworkbook = new HSSFWorkbook(is);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
            result.setSuccess(false);
            throw new ValidateRuntimeException(CommonUtil.getStackTrace(e));
        }

        for (VSheet sheet : rules) {
            List<ValidateRule> coumnRules = sheet.getColumns();
            HashMap<Integer, ValidateRule> coumnRules_Map = new HashMap<Integer, ValidateRule>();
            List<UniqueKey> uniqueKeys = sheet.getUniqueKeys();
            List<Object> rowKeys = new ArrayList();
            for (ValidateRule columnRow : coumnRules) {
                rowKeys.add(columnRow.getColumnIndex());
                coumnRules_Map.put(new Integer(columnRow.getColumnIndex()), columnRow);
            }



            HSSFSheet hssfsheet = hssfworkbook.getSheetAt(sheet.getSheetIndex());
            int rows = hssfsheet.getLastRowNum();

            HashMap<String, Integer> countIdt = new HashMap(rows * uniqueKeys.size() + 10, 1F);

            int endRow = sheet.getEndRow();
            if (sheet.getEndRow() != null && hssfsheet.getLastRowNum() > endRow) {
                result.setSuccess(false);
                result.getErrorMsg().append("解析工作表失败:表格sheet数据不能超过" + sheet.getEndRow() + "条"+"");

            }
            excelCounts += (hssfsheet.getLastRowNum() - sheet.getBeginRow() + 1);
            try {

                for (int rowNum = sheet.getBeginRow(); rowNum <= hssfsheet.getLastRowNum(); rowNum++) {

                    HSSFRow hssfRow = hssfsheet.getRow(rowNum);
                     if(hssfRow==null){
                         excelCounts--;
                        continue;
                    }
                    Boolean empty =  checkHSSFRowIsEmpty(hssfRow);
                     if(empty){
                         excelCounts--;
                         continue;
                     }

                    for (Object key : rowKeys) {
                        if (hssfRow.getCell((Integer) key) == null) {
                            hssfRow.createCell((Integer) key);
                            hssfRow.getCell((Integer) key).setCellType(Cell.CELL_TYPE_STRING);
                            hssfRow.getCell((Integer) key).setCellValue("");
                        }
                        Cell cell = hssfRow.getCell((Integer) key);
                        String cellText = getCellText(cell);

                        Message msg = RuleEngine.process(cellText, coumnRules_Map.get(key));
                        if (!msg.isSuccess()) {
                            result.setSuccess(false);
                            result.getErrorMsg().append("第" + (rowNum + 1) + "行:" + msg.getMsg()+"");

                        }

                       }

                    for (UniqueKey uniqueRule : uniqueKeys) {
                        List<Integer> keyRows = uniqueRule.getUniqueColumn();
                        String keyString = uniqueRule.getKeyName();
                        for (Integer key : keyRows) {
                            if (hssfRow.getCell((Integer) key) == null) {
                                hssfRow.createCell((Integer) key);
                                hssfRow.getCell((Integer) key).setCellType(Cell.CELL_TYPE_STRING);
                                hssfRow.getCell((Integer) key).setCellValue("");
                            }
                            Cell cell = hssfRow.getCell((Integer) key);
                            String cellText = getCellText(cell);
                            if (CommonUtil.isNull(cellText)) {
                                keyString = "";
                                break;
                            }
                            keyString += "--" + cellText;
                        }

                        if (!CommonUtil.isNull(keyString)) {
                            if (countIdt.containsKey(keyString)) {
                                result.setSuccess(false);
                                result.getErrorMsg().append("第" + (rowNum + 1) + "行:" + "唯一性约束不通过，" + keyString + "表格内已存在"+"");

                            } else {
                                countIdt.put(keyString, new Integer(1));
                            }
                        }

                    }
                    count++;
                }
            } catch (Exception e) {
                e.printStackTrace();
                closeInStream(is);
               throw new ValidateXmlException(CommonUtil.getStackTrace(e));
            }

        }

        closeInStream(is);
        if (count == excelCounts.intValue()||result.getSuccess() && count != 0)
            result.setSuccess(true);
        else {

            result.setSuccess(false);

        }

        return result;
    }

    private static void closeInStream(InputStream is){
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                // TODO Auto-generated catch block
                throw new ValidateXmlException(CommonUtil.getStackTrace(e));
            }
        }
    }

    public static ValidateResult readExcel_XLSX(InputStream is, List<VSheet> rules,String excelType) {
        Integer excelCounts = 0;
        int count = 0;
        ValidateResult result = new ValidateResult();
        result.setSuccess(true);
        StringBuilder msgs = new StringBuilder();
        result.setErrorMsg(msgs);

        XSSFWorkbook xssfworkbook = null;
        try {
            xssfworkbook = new XSSFWorkbook(is);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
            throw new ValidateRuntimeException("解析工作表失败"+e.toString());
        }

        for (VSheet sheet : rules) {
            List<ValidateRule> coumnRules = sheet.getColumns();
            HashMap<Integer, ValidateRule> coumnRules_Map = new HashMap<Integer, ValidateRule>();
            List<UniqueKey> uniqueKeys = sheet.getUniqueKeys();
            List<Object> rowKeys = new ArrayList();
            for (ValidateRule columnRow : coumnRules) {
                rowKeys.add(columnRow.getColumnIndex());
                coumnRules_Map.put(new Integer(columnRow.getColumnIndex()), columnRow);
            }


            XSSFSheet xssfsheet = xssfworkbook.getSheetAt(sheet.getSheetIndex());
            int rows = xssfsheet.getLastRowNum();

            HashMap<String, Integer> countIdt = new HashMap(rows * uniqueKeys.size() + 10, 1F);

            int endRow = sheet.getEndRow();
            if (sheet.getEndRow() != null && xssfsheet.getLastRowNum() > endRow) {
                result.setSuccess(false);
                result.getErrorMsg().append("解析工作表失败:表格sheet数据不能超过" + sheet.getEndRow() + "条"+"");

            }
            excelCounts += (xssfsheet.getLastRowNum() - sheet.getBeginRow() + 1);
            try {

                for (int rowNum = sheet.getBeginRow(); rowNum <= xssfsheet.getLastRowNum(); rowNum++) {
                    XSSFRow xssfRow = xssfsheet.getRow(rowNum);
                    if(xssfRow==null){
                        excelCounts--;
                        continue;
                    }
                    Boolean empty =  checkXSSFRowIsEmpty(xssfRow);
                    if(empty){
                        excelCounts--;
                        continue;
                    }
                    for (Object key : rowKeys) {
                        if (xssfRow.getCell((Integer) key) == null) {
                            xssfRow.createCell((Integer) key);
                            xssfRow.getCell((Integer) key).setCellType(Cell.CELL_TYPE_STRING);
                            xssfRow.getCell((Integer) key).setCellValue("");
                        }
                        Cell cell = xssfRow.getCell((Integer) key);
                        String cellText = getXssCellText(cell);

                        Message msg = RuleEngine.process(cellText, coumnRules_Map.get(key));
                        if (!msg.isSuccess()) {
                            result.setSuccess(false);
                            result.getErrorMsg().append("第" + (rowNum + 1) + "行:" + msg.getMsg()+"");

                        }

                       }

                    for (UniqueKey uniqueRule : uniqueKeys) {
                        List<Integer> keyRows = uniqueRule.getUniqueColumn();
                        String keyString = uniqueRule.getKeyName();
                        for (Integer key : keyRows) {
                            if (xssfRow.getCell((Integer) key) == null) {
                                xssfRow.createCell((Integer) key);
                                xssfRow.getCell((Integer) key).setCellType(Cell.CELL_TYPE_STRING);
                                xssfRow.getCell((Integer) key).setCellValue("");
                            }
                            Cell cell = xssfRow.getCell((Integer) key);
                            String cellText = getXssCellText(cell);
                            if (CommonUtil.isNull(cellText)) {
                                keyString = "";
                                break;
                            }
                            keyString += "--" + cellText;
                        }

                        if (!CommonUtil.isNull(keyString)) {
                            if (countIdt.containsKey(keyString)) {
                                result.setSuccess(false);
                                result.getErrorMsg().append("第" + (rowNum + 1) + "行:" + "唯一性约束不通过，" + keyString + "表格内已存在"+"");

                            } else {
                                countIdt.put(keyString, new Integer(1));
                            }
                        }

                    }
                    count++;
                }
            } catch (Exception e) {
                e.printStackTrace();
                closeInStream(is);
                throw new ValidateXmlException(CommonUtil.getStackTrace(e));
            }

        }

        closeInStream(is);
        if ( count == excelCounts.intValue()&&result.getSuccess() && count != 0)
            result.setSuccess(true);
        else {

            result.setSuccess(false);
        }

        return result;
    }

}
