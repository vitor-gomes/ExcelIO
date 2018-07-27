/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package main;

import com.io.excel.ExcelDocument;
import com.io.excel.FileType;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author Vitor
 */
public class Main {
    
    public static void main(String[] args) throws ClassNotFoundException, IOException {
        
        List<Pessoa> pessoas = new ArrayList<Pessoa>();
        pessoas.add(new Pessoa(10, "Fulano", "Fulano"));
        pessoas.add(new Pessoa(10, "Ciclano", "Ciclano"));
        pessoas.add(new Pessoa(10, "Beltrano", "Beltrano"));
        pessoas.add(new Pessoa(10, "Fulano", "Fulano"));
        
        ExcelDocument doc = new ExcelDocument("arquivo-teste4", FileType.XLSX);
        doc.addSheet("pessoas", pessoas);
        
        doc.getOutputStream();
        
        
    }
    
}
