package com.miluo.data.util;

import com.miluo.data.dto.FieldDTO;
import com.miluo.data.dto.TableDTO;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtil {

    public static void getDictionary(List<TableDTO> tables) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        HSSFSheet qd = wb.createSheet("表清单");
        setQingDan(qd,tables,wb,createHelper);

        for (TableDTO table:tables) {
            System.out.println(table.getTableName());
            String tableName = table.getTableName().substring(0,3).toUpperCase()+"_"+table.getTableRemark();
            System.out.println(tableName);
            HSSFSheet sheet = wb.createSheet(tableName);
            setColumnStyle(sheet);
            setExcelHeader(sheet,table);
            setExcelBody(sheet,table.getFields(),createHelper,wb);
            wb.setForceFormulaRecalculation(true);
        }

        String exportPath = "D:\\temp.xls";
        FileOutputStream output=new FileOutputStream(exportPath);
        wb.write(output);
        output.flush();
        System.out.println("excel success!");
    }


    private static void setExcelHeader(HSSFSheet sheet,TableDTO table){
        String tableName = table.getTableName().substring(0,3).toUpperCase()+"_"+table.getTableRemark();
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("表:"+tableName);

        HSSFRow row1 =sheet.createRow(1);
        HSSFCell cell11 = row1.createCell(0);
        cell11.setCellValue("名称");
        HSSFCell cell12 = row1.createCell(2);
        cell12.setCellValue(tableName);

        HSSFRow row2 =sheet.createRow(2);
        HSSFCell cell21 = row2.createCell(0);
        cell21.setCellValue("代码");
        HSSFCell cell22 = row2.createCell(2);
        cell22.setCellValue(table.getTableName());

        HSSFRow row4 =sheet.createRow(4);
        HSSFCell cell41 = row4.createCell(0);
        cell41.setCellValue(tableName);
    }

    private static void setExcelBody(HSSFSheet sheet,List<FieldDTO> fields,CreationHelper createHelper,HSSFWorkbook wb){
        HSSFRow row = sheet.createRow(5);
        HSSFCell cell1 = row.createCell(0);
        cell1.setCellValue("名称");
        HSSFCell cell2 = row.createCell(1);
        cell2.setCellValue("代码");
        HSSFCell cell3 = row.createCell(2);
        cell3.setCellValue("数据类型");
        HSSFCell cell4 = row.createCell(3);
        cell4.setCellValue("长度");
        HSSFCell cell5 = row.createCell(4);
        cell5.setCellValue("注释");
        HSSFCell cell6 = row.createCell(5);
        cell6.setCellValue("PK");
        HSSFCell cell7 = row.createCell(6);
        cell7.setCellValue("FK");
        HSSFCell cell8 = row.createCell(7);
        cell8.setCellValue("UK");

        int rowNum = 6;
        for (FieldDTO field:fields) {
            HSSFRow dataRow = sheet.createRow(rowNum);
            HSSFCell dataCell1 = dataRow.createCell(0);
            dataCell1.setCellValue(field.getColumnName());
            HSSFCell dataCell2 = dataRow.createCell(1);
            dataCell2.setCellValue(field.getColumnCode());
            HSSFCell dataCell3 = dataRow.createCell(2);
            dataCell3.setCellValue(field.getDataType());
            HSSFCell dataCell4 = dataRow.createCell(3);
            dataCell4.setCellValue(field.getSize());
            HSSFCell dataCell5 = dataRow.createCell(4);
            dataCell5.setCellValue(field.getComment());
            HSSFCell dataCell6 = dataRow.createCell(5);
            dataCell6.setCellValue(field.getPK());
            HSSFCell dataCell7 = dataRow.createCell(6);
            dataCell7.setCellValue(field.getFK());
            HSSFCell dataCell8 = dataRow.createCell(7);
            dataCell8.setCellValue(field.getUK());
            rowNum++;
        }


        CellStyle cellStyle = setFontBule(wb);
        HSSFRow backRow = sheet.createRow(++rowNum);
        HSSFCell back = backRow.createCell(0);
        back.setCellValue("返回表清单");
        back.setCellStyle(cellStyle);
        Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        hyperlink.setAddress("'表清单'!A1");
        back.setHyperlink(hyperlink);

    }

    private static void setQingDan(HSSFSheet sheet,List<TableDTO> tables,HSSFWorkbook wb,CreationHelper createHelper){
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 3);
        sheet.addMergedRegion(region);
        sheet.setColumnWidth(0, 256*45);
        sheet.setColumnWidth(1, 256*45);
        sheet.setColumnWidth(2, 256*45);
        sheet.setColumnWidth(3, 256*60);
        HSSFRow row1 = sheet.createRow(0);
        HSSFCell cell = row1.createCell(0);
        cell.setCellValue("表清单");
        cell.setCellStyle(setFontStyle(wb,512));
        row1.setHeight((short) 600);
        HSSFRow info = sheet.createRow(1);
        info.setHeight((short) 400);
        HSSFCell cell3 = info.createCell(0);
        cell3.setCellStyle(setFontStyle(wb,256));
        cell3.setCellValue("名称");
        HSSFCell cell4 = info.createCell(1);
        cell4.setCellStyle(setFontStyle(wb,256));
        cell4.setCellValue("表名");
        HSSFCell cell5 = info.createCell(2);
        cell5.setCellStyle(setFontStyle(wb,256));
        cell5.setCellValue("前缀缩写");
        HSSFCell cell6 = info.createCell(3);
        cell6.setCellStyle(setFontStyle(wb,256));
        cell6.setCellValue("表类型(事实（FAC），维度DIM，汇总AGS)");

        int rowNum = 2;
        CellStyle cellStyle = setFontBule(wb);
        for (TableDTO table:tables) {
            HSSFRow row = sheet.createRow(rowNum);
            row.setHeight((short) 300);
            String tableName = table.getTableName().substring(0,3).toUpperCase()+"_"+table.getTableRemark();

            HSSFCell cell1 = row.createCell(0);
            cell1.setCellValue(tableName);
            cell1.setCellStyle(cellStyle);
            Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
            String address = "'"+tableName+"'"+"!A1";
            hyperlink.setAddress(address);
            cell1.setHyperlink(hyperlink);

            HSSFCell cell2 = row.createCell(1);
            cell2.setCellValue(table.getTableName());
            cell2.setCellStyle(cellStyle);
            Hyperlink hyperlink2 = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
            hyperlink2.setAddress(address);
            cell2.setHyperlink(hyperlink2);

            rowNum++;
        }
    }



    private static void setColumnStyle(HSSFSheet sheet){
        sheet.setColumnWidth(0, 256*25);
        sheet.setColumnWidth(1, 256*25);
        sheet.setColumnWidth(2, 256*25);
        sheet.setColumnWidth(3, 256*25);
        sheet.setColumnWidth(4, 256*25);
    }

    private static CellStyle setFontBule(HSSFWorkbook wb){
        // 设置字体
        CellStyle redStyle = wb.createCellStyle();
        HSSFFont redFont = wb.createFont();
        //颜色
        redFont.setColor((short) 30);
        redFont.setUnderline((byte) 1);
        //字体
        redStyle.setFont(redFont);
        return redStyle;
    }


    private static CellStyle setFontStyle(HSSFWorkbook wb,int size){
        // 设置字体
        CellStyle redStyle = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setFontHeight((short) size);
        redStyle.setAlignment(HorizontalAlignment.CENTER);
        redStyle.setFont(font);
        redStyle.setTopBorderColor((short) 128);
        redStyle.setBottomBorderColor((short) 128);
        redStyle.setLeftBorderColor((short) 128);
        redStyle.setRightBorderColor((short) 128);
        return redStyle;
    }
}
