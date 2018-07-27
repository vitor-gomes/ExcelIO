/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.io.excel;

import java.sql.ResultSet;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Vitor
 */
public class XLSXBuilder extends Builder {
    
    private XSSFWorkbook xssfw;

    @Override
    public void createWorkbook() {
        this.xssfw = new XSSFWorkbook();
    }

    @Override
    public void addSheet(String sheetName, ResultSet rs) {
        SXSSFWorkbook sxssfw = new SXSSFWorkbook(xssfw);
        SXSSFSheet sheet = sxssfw.createSheet(sheetName);
        sheet.setRandomAccessWindowSize(150);
        try {
            super.writeOnSheet(rs, xssfw, sheet);
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            sxssfw.dispose();
        }
    }

    @Override
    public <T> void addSheet(String sheetName, List<T> list) {
        SXSSFWorkbook sxssfw = new SXSSFWorkbook(xssfw);
        SXSSFSheet sheet = sxssfw.createSheet(sheetName);
        sheet.setRandomAccessWindowSize(150);
        try {
            super.writeOnSheet(list, xssfw, sheet);
        } catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            sxssfw.dispose();
        }
    }

    @Override
    public Workbook getWorkbook() {
        return xssfw;
    }
}
