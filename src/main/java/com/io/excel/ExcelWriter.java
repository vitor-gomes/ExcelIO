/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.io.excel;

import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Timestamp;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import javax.persistence.Column;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;

/**
 *
 * @author Vitor
 */
public class ExcelWriter<O> {

    private Class<O> c;

    public ExcelWriter(Class<O> c) {
        this.c = c;
    }

    public ByteArrayOutputStream getXLS(List<O> list, String nomeArquivo) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        return getDocument(list, workbook, workbook.createSheet(nomeArquivo));
    }

    public ByteArrayOutputStream getXLSX(List<O> list, String nomeArquivo) throws Exception {
        XSSFWorkbook xssfw = new XSSFWorkbook();
        SXSSFWorkbook sxssfw = new SXSSFWorkbook(xssfw);

        SXSSFSheet sheet = sxssfw.createSheet(nomeArquivo);
        sheet.setRandomAccessWindowSize(150);
        try {
            return getDocument(list, sxssfw, sheet);
        } finally {
            sxssfw.dispose();
        }
    }

    public ByteArrayOutputStream getXLSX(ResultSet rs, String nomeArquivo) throws Exception {
        XSSFWorkbook xssfw = new XSSFWorkbook();
        SXSSFWorkbook sxssfw = new SXSSFWorkbook(xssfw);
        SXSSFSheet sheet = sxssfw.createSheet(nomeArquivo);
        sheet.setRandomAccessWindowSize(150);
        try {
            return getDocument(rs, sxssfw, sheet);
        } finally {
            sxssfw.dispose();
        }
    }

    private ByteArrayOutputStream getDocument(List<O> list, Workbook workbook, Sheet sheet) throws Exception {
        if (list == null || list.isEmpty()) {
            return null;
        }
        Field[] fields = c.getDeclaredFields();
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
            } catch (java.lang.NoClassDefFoundError e) {
                cell.setCellValue(f.getName());
            }
        }

        NumberFormat numberFormat = NumberFormat.getInstance(new Locale("pt", "BR"));
        for (O obj : list) {
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
                            break;
                        case "java.lang.Integer":
                            c.setCellValue(numberFormat.format(o));
                            break;
                        case "java.lang.Double":
                            c.setCellValue(numberFormat.format(o));
                            break;
                        default:
                            c.setCellValue(o.toString());
                            break;
                    }
                }
            }
        }

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbook.write(baos);
        baos.close();
        return baos;
    }

    private ByteArrayOutputStream getDocument(ResultSet rs, Workbook workbook, Sheet sheet) throws Exception {
        int rownum = 0;
        int cellnum = 0;
        Row row = sheet.createRow(rownum++);

        CellStyle header = workbook.createCellStyle();
        header.setBorderBottom(BorderStyle.THIN);
        header.setBorderLeft(BorderStyle.THIN);
        header.setBorderRight(BorderStyle.THIN);
        header.setBorderTop(BorderStyle.THIN);
        header.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        header.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        ResultSetMetaData rsmd = rs.getMetaData();
        for (int i = 1; i <= rsmd.getColumnCount(); i++) {
            Cell cell = row.createCell(cellnum++);
            cell.setCellValue(rsmd.getColumnName(i));
            cell.setCellStyle(header);
        }

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
                        Object o = rs.getObject(i);
                        cell.setCellValue(o.toString());
                        break;
                }
            }
        }

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbook.write(baos);
        baos.close();
        return baos;
    }
}
