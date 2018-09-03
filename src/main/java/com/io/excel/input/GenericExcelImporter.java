package com.io.excel.input;

import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author Pedro Coelho
 */
public class GenericExcelImporter extends AbstractExcelImporter {

    // TODO: 2018-07-27 - implement this use case.
    public GenericExcelImporter(Object fileObject, String dirPath, int headersize) {
        super(fileObject, dirPath, headersize);
    }

    @Override
    public boolean trataRow(Row r, int lineNo) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean processaRows() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
    
}
