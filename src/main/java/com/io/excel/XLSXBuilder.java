/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.io.excel;

import java.sql.ResultSet;
import java.util.List;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
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
            super.writeOnSheet(list, sxssfw, sxssfw.createSheet(sheetName));
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
