/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.io.excel;

import java.sql.ResultSet;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author Vitor
 */
class XLSBuilder extends Builder {
    
    private HSSFWorkbook hssfw;

    @Override
    public void createWorkbook() {
        this.hssfw = new HSSFWorkbook();
    }

    @Override
    public void addSheet(String sheetName, ResultSet rs) {
        try {
            super.writeOnSheet(rs, hssfw, hssfw.createSheet(sheetName));
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    @Override
    public <T> void addSheet(String sheetName, List<T> list) {
        try {
            super.writeOnSheet(list, hssfw, hssfw.createSheet(sheetName));
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    @Override
    public Workbook getWorkbook() {
        return hssfw;
    }
}
