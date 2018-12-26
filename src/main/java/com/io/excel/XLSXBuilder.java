/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.io.excel;

import java.lang.reflect.Field;
import java.sql.ResultSet;
import java.util.List;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IgnoredErrorType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Vitor
 */
class XLSXBuilder extends Builder {
    
    private SXSSFWorkbook sxssfw;

    @Override
    public void createWorkbook() {
        this.sxssfw = new SXSSFWorkbook(new XSSFWorkbook());
    }

    @Override
    public void addSheet(String sheetName, ResultSet rs) {
        try {
            super.writeOnSheet(rs, sxssfw, sxssfw.createSheet(sheetName));
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    @Override
    public <T> void addSheet(String sheetName, List<T> list) {
        try {
            SXSSFSheet sheet = sxssfw.createSheet(sheetName);
            
            Field _sh = SXSSFSheet.class.getDeclaredField("_sh");
            _sh.setAccessible(true);
            XSSFSheet xssfsheet = (XSSFSheet)_sh.get(sheet);
            xssfsheet.addIgnoredErrors(new CellRangeAddress(0,1048575,0,9999),IgnoredErrorType.NUMBER_STORED_AS_TEXT );
            
            // Adding tracking for autosizable columns for XSSF Workbooks.
            if(super.isAutosizeAll() || super.isAutosizeHeader()) {
                sheet.trackAllColumnsForAutoSizing();
            } else if (!super.autosizableColumns.isEmpty()) {
                sheet.trackColumnsForAutoSizing(super.autosizableColumns);
            }
            
            super.writeOnSheet(list, sxssfw, sheet);
        } catch (Exception ex) {
            ex.printStackTrace();
        } 
    }
    
    @Override
    public Workbook getWorkbook() {
        return sxssfw;
    }

    @Override
    public CellStyle getHeader() {
        if(super.header == null) {
            super.header = super.defaultHeader(sxssfw);
        }
        return super.header;
    }

    @Override
    public Font getFont() {
        return sxssfw.createFont();
    }
}
