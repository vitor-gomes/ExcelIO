/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.io.excel;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.ResultSet;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

/**
 *
 * @author Vitor
 */
public class ExcelDocument {
    
    private String nome;
    private FileType type;
    private Builder builder;
    
    public ExcelDocument(String nome, FileType type) throws ClassNotFoundException {
        this.nome = nome;
        this.type = type;
        switch(type) {
            case XLS:
                builder = new XLSBuilder();
                break;
            case XLSX:
                builder = new XLSXBuilder();
                break;
            default:
                throw new ClassNotFoundException("Type of file not found.");
        }
        builder.createWorkbook();
    }
    
    public void addSheet(String sheetName, ResultSet rs) {
        builder.addSheet(sheetName, rs);
    }
    
    public <T> void addSheet(String sheetName, List<T> list) {
        builder.addSheet(sheetName, list);
    }
    
    public CellStyle getHeader() {
        return builder.getHeader();
    }
    
    public Font getFont() {
        return builder.getFont();
    }
    
    /**
     * WARNING: This can may result in great slowdown for large sheets!
     * <p>
     * Autosizes all columns.
     */
    public void autosizeAll() {
        builder.setAutosizeAll(true);
    }
    
    /**
     * WARNING: This can may result in great slowdown for large sheets!
     * <p>
     * Sets which columns the engine should track to Autosize.
     * <p>
     * @param autosizableColumns - a list of the indices of all columns to track.
     */
    public void setAutosizableColumns(List<Integer> autosizableColumns) {
        builder.setAutosizableColumns(autosizableColumns);
    }

    /**
     * WARNING: This can may result in great slowdown for large sheets!
     * <p>
     * Sets column to track for Autosize.
     * <p>
     * @param column - index of column to track.
     */
    public void addAutosizableColumn(int column) {
        builder.addAutosizableColumn(column);
    }
    
    /**
     * Sets the widths (in pixels) for the desired columns
     * <p>
     * @param columnsWidth - Set with values for the width keys.
     */
    public void setColumnsWidth(Map<Integer, Integer> columnsWidth) {
        builder.setColumnsWidth(columnsWidth);
    }
    
    /**
     * Sets the width (in pixels) of de desired column.
     * <p>
     * @param column
     * @param width 
     */
    public void addColumnWidth(int column, int width) {
        builder.addColumnWidth(column, width);
    }
    
    /**
     * Sets the height (in points) of the header.<p>
     * Fallback to engine default if set to 0.
     * 
     * @param height 
     */
    public void setHeaderHeight(int height) {
        builder.setHeaderHeight(height);
    }
    
    public ByteArrayOutputStream getByteArrayOutputStream() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        builder.getWorkbook().write(baos);
        baos.close();
        return baos;
    }
    
    public OutputStream getOutputStream() throws FileNotFoundException, IOException {
        OutputStream outputStream = new FileOutputStream(nome + type.getExtension());
        this.getByteArrayOutputStream().writeTo(outputStream);   
        return outputStream;
    }

    public String getNome() {
        return nome;
    }

    public void setNome(String nome) {
        this.nome = nome;
    }

    public FileType getType() {
        return type;
    }

    public void setType(FileType type) {
        this.type = type;
    }
    
}
