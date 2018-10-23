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
