/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.io.excel;

import com.io.excel.utils.PixelUtil;
import java.lang.reflect.Field;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import javax.persistence.Column;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 *
 * @author Vitor
 */
abstract class Builder {
    
    protected CellStyle header;
    
    protected Boolean autosizeAll = false;
    protected Boolean autosizeHeader = false;
    protected List<Integer> autosizableColumns = new ArrayList();
    protected Map<Integer, Integer> columnsWidth = new HashMap();
    protected Integer headerHeight = 0; // 0 == default
    
    public abstract void createWorkbook();
    public abstract Workbook getWorkbook();
    
    public abstract Font getFont();
    public abstract CellStyle getHeader();
    
    public abstract void addSheet(String sheetName, ResultSet rs);    
    public abstract <T> void addSheet(String sheetName, List<T> list);

    
    protected CellStyle defaultHeader(Workbook workbook) {
        CellStyle defaultHeader = workbook.createCellStyle();
        defaultHeader.setBorderBottom(BorderStyle.THIN);
        defaultHeader.setBorderLeft(BorderStyle.THIN);
        defaultHeader.setBorderRight(BorderStyle.THIN);
        defaultHeader.setBorderTop(BorderStyle.THIN);
        defaultHeader.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        defaultHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return defaultHeader;
    }
        
    protected void writeOnSheet(ResultSet rs, Workbook workbook, Sheet sheet) throws Exception {
        int rownum = 0;
        int cellnum = 0;
        
        Row row = sheet.createRow(rownum++);
        
        if(header == null) {
            header = defaultHeader(workbook);
        }
        
        ResultSetMetaData rsmd = rs.getMetaData();
        for (int i = 1; i <= rsmd.getColumnCount(); i++) {
            Cell cell = row.createCell(cellnum++);
            cell.setCellValue(rsmd.getColumnName(i));
            cell.setCellStyle(header);
        }

        formatHeader(sheet);
        
        DataFormat fmt = workbook.createDataFormat();

        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(fmt.getFormat("m/d/yy"));

        CellStyle doubleStyle = workbook.createCellStyle();
        doubleStyle.setDataFormat(fmt.getFormat("0.00"));

        while (rs.next()) {
            Row r = sheet.createRow(rownum++);
            cellnum = 0;
            for (int i = 1; i <= rsmd.getColumnCount(); i++) {
                Cell cell = r.createCell(cellnum++);
                Object object = rs.getObject(i);

                if (object == null) {
                    cell.setCellType(CellType.BLANK);
                    continue;
                }
                switch (rsmd.getColumnClassName(i)) {
                    case "java.lang.Integer":
                        Integer integer = (Integer) object;
                        cell.setCellValue(integer);
                        break;
                    case "java.math.BigDecimal":
                        BigDecimal bd = (BigDecimal) object;
                        cell.setCellValue(bd.doubleValue());
                        cell.setCellStyle(doubleStyle);
                        break;
                    case "java.sql.Timestamp":
                        Timestamp t = (Timestamp) object;
                        Date date = new Date(t.getTime());
                        cell.setCellValue(date);
                        cell.setCellStyle(dateStyle);
                        break;
                    default:
                        if(!(object instanceof java.util.Collection)) {
                            Object o = rs.getObject(i);
                            cell.setCellValue(o.toString());
                        }
                }
            }
        }
        
        if (autosizeAll)
            autosize(sheet);
    }
    
    protected <T> void writeOnSheet(List<T> list, Workbook workbook, Sheet sheet) throws Exception {
        if (list == null || list.isEmpty()) {
            return;
        }        
        if(header == null) {
            header = defaultHeader(workbook);
        }
        Field[] fields = list.get(0).getClass().getDeclaredFields();
        int rownum = 0, cellnum = 0;
        Row row = sheet.createRow(rownum++);
        
        for (Field f : fields) {
            Cell cell = row.createCell(cellnum++);
            try {
                Column column = f.getAnnotation(Column.class);
                if (column == null) {
                    cell.setCellValue(f.getName());
                } else {
                    cell.setCellValue(column.name());
                }
                cell.setCellStyle(header);
            } catch (java.lang.NoClassDefFoundError e) {
                cell.setCellValue(f.getName());
            }
        }

        formatHeader(sheet);
        
        DataFormat fmt = workbook.createDataFormat();
        
        CellStyle textStyle = workbook.createCellStyle();
        textStyle.setDataFormat(fmt.getFormat("@"));
        
        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(fmt.getFormat("m/d/yy"));

        CellStyle doubleStyle = workbook.createCellStyle();
        doubleStyle.setDataFormat(fmt.getFormat("0.00"));
        
        for (T obj : list) {
            cellnum = 0;
            Row r = sheet.createRow(rownum++);

            for (Field f : fields) {
                Cell c = r.createCell(cellnum++);
                f.setAccessible(true);
                Object o = f.get(obj);
                Type t = f.getGenericType();

                if (o == null) {
                    c.setCellType(CellType.BLANK);
                } else {
                    switch (t.getTypeName()) {
                        case "java.util.Date":
                            Column column = f.getAnnotation(Column.class);
                            String pattern = column.columnDefinition();
                            SimpleDateFormat frmt = new SimpleDateFormat(pattern);
                            c.setCellValue(frmt.format(o));
                            c.setCellStyle(dateStyle);
                            break;
                        case "java.lang.Integer":
                            Integer i = (Integer) o;
                            c.setCellValue(i);
                            break;
                        case "java.lang.Double":
                            Double d = (Double) o;
                            c.setCellValue(d);
                            c.setCellStyle(doubleStyle);
                            break;
                        default:
                            c.setCellValue(o.toString());
                            c.setCellStyle(textStyle);
                            break;
                    }
                }
            }
        }
        
        if (autosizeAll)
            autosize(sheet);
    }
    
    private void autosize(Sheet sheet) {
        for (int c : autosizableColumns) 
                sheet.autoSizeColumn(c);
    }
    
    private void formatHeader(Sheet sheet) {
        Row row = sheet.getRow(0);
        if(headerHeight != 0) 
            row.setHeightInPoints(headerHeight);
        
        if (autosizeHeader || autosizeAll) {
            int c = row.getPhysicalNumberOfCells();
            for (int i = 0; i < c; i++) 
                autosizableColumns.add(i);
        } 
        if (autosizeHeader) {
            autosize(sheet);
        }
        if (!autosizeAll && sheet instanceof SXSSFSheet) {
            SXSSFSheet sXSSFSheet = (SXSSFSheet) sheet;
            sXSSFSheet.untrackAllColumnsForAutoSizing();
        }
        
        if (!columnsWidth.isEmpty()) {
            for (Map.Entry<Integer, Integer> columnWidth : columnsWidth.entrySet()) 
                sheet.setColumnWidth(columnWidth.getKey(), PixelUtil.pixel2WidthUnits(columnWidth.getValue()));
        }
    }

    public void setAutosizeHeader(boolean autosizeHeader) {
        this.autosizeHeader = autosizeHeader;
    }
    
    public boolean isAutosizeHeader() {
        return autosizeHeader;
    }
    
    public boolean isAutosizeAll() {
        return autosizeAll;
    }

    public void setAutosizeAll(boolean autosizeAll) {
        this.autosizeAll = autosizeAll;
    }

    public List<Integer> getAutosizableColumns() {
        return autosizableColumns;
    }

    public void setAutosizableColumns(List<Integer> autosizableColumns) {
        this.autosizableColumns = autosizableColumns;
    }

    public void addAutosizableColumn(int column) {
        this.autosizableColumns.add(column);
    }
    
    public Map<Integer, Integer> getColumnsWidth() {
        return columnsWidth;
    }

    public void setColumnsWidth(Map<Integer, Integer> columnsWidth) {
        this.columnsWidth = columnsWidth;
    }
    
    public void addColumnWidth(int column, int width) {
        this.columnsWidth.put(column, width);
    }

    public Integer getHeaderHeight() {
        return headerHeight;
    }

    public void setHeaderHeight(Integer headerHeight) {
        this.headerHeight = headerHeight;
    }
    
}
