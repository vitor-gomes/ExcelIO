package com.io.excel.input;

import com.monitorjbl.xlsx.StreamingReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import org.apache.commons.fileupload.FileItem;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author Pedro Coelho
 */
public abstract class AbstractExcelImporter {
    
    /**
     * m_list deve ser utilizada como uma lista simples
     * onde a key Integer deve ser o número da linha
     */
    protected Map<Integer, Object> m_list = new TreeMap();
    /**
     * m_map deve ser utilizada como um map, onde, por exemplo, a importação 
     * da base deve acumular valores presentes em múltiplas linhas.
     */
    protected Map<Object, Object> m_map = new HashMap();
    
    protected int headerSize;
    protected static String dirPath;
    
    private static final int F_ROWCACHE = 100;
    private static final int F_BUFFERSIZE = 2048;
    private int rowCache;
    private int bufferSize;
    private boolean singleSheet = true;
    private boolean firstSheetOnly = false;
    private boolean header = true;
    
    private File file;
    private Object fileObject;
    private FileItem fileItem;
    private InputStream uploadedStream;
    protected boolean success;
    protected boolean success100percent = true;
    protected List<String> errors = new ArrayList<>();
    private TYPE type;
    
    public enum TYPE {
        HSSF_WORKBOOK, XSSF_WORKBOOK, INVALID
    }
    
    public AbstractExcelImporter(Object fileObject, String dirPath, int headersize) {
        this(fileObject, dirPath, headersize, F_ROWCACHE, F_BUFFERSIZE);
    }
    
    public AbstractExcelImporter(Object fileObject, String dirPath, int headersize, int bufferSize, int rowCache) {
        this.rowCache = rowCache;
        this.bufferSize = bufferSize;
        this.headerSize = headersize;
        this.fileObject = fileObject;
        this.dirPath = dirPath;
        try {
            realizaUpload();
            success = true;
        } catch (Exception e) {
            e.printStackTrace();
            success = false;
            errors.add(e.getMessage());
        }
    }
    
    private void realizaUpload() throws Exception {
        fileItem = (FileItem) fileObject;
        String fileName = fileItem.getName();
        
        fileName = fileName.contains("\\") ? fileName.substring(fileName.lastIndexOf("\\") + 1) : fileName;
        
        if (fileName.length() > 0) {
            try{
                String path = dirPath + File.separator + fileName;

                uploadedStream = fileItem.getInputStream();

                type = defineType(uploadedStream);

                switch(type) {
                    case INVALID:
                        //TODO: set errors string elsewhere!!!!
                        throw(new Exception("Invalid file!"));
                    case HSSF_WORKBOOK:
                        escreveXLS(path);
                    case XSSF_WORKBOOK:
                        escreveXLSX(path);
                        break;
                }
            
            } catch (Exception e) { throw e; }
            finally { uploadedStream.close(); }
            
        } else {
            throw(new Exception("Arquivo Inexistente!"));
        }
        
    }
    
    /**
     * Método de escrita do arquivo em disco.
     * <p>
     * Este método escreve o arquivo XLS em disco.
     * Dar override neste método da implementação da classe quando desejar bloquear
     * o upload de arquivos XLS
     * <p>
     * @param path caminho completo do arquivo a ser escrito
     * @throws Exception 
     */
    protected void escreveXLS(String path) throws Exception {
        escreve(path);
    }
    
    /**
     * Método de escrita do arquivo em disco.
     * <p>
     * Este método escreve o arquivo XLSX em disco.
     * Dar override neste método da implementação da classe quando desejar bloquear
     * o upload de arquivos XLSX
     * <p>
     * @param path caminho completo do arquivo a ser escrito
     * @throws Exception 
     */
    protected void escreveXLSX(String path) throws Exception {
        escreve(path);
    }
    
    private void escreve(String path) throws Exception {
        
        file = new File(path);
        final int TAM_BUFF = (8 * 1024);
        byte arr[] = new byte[TAM_BUFF];
        int bytesLidos;
        try (OutputStream out = new FileOutputStream(file)) {
            uploadedStream = fileItem.getInputStream();
            while ((bytesLidos = uploadedStream.read(arr, 0, TAM_BUFF)) >= 0) {
                out.write(arr, 0, bytesLidos);
            }
        } catch (Exception e) { throw e; }
        
    }
    
    /**
     * Método de importação dos dados do Excel
     * <p>
     * Método que deve ser chamado caso após a chamada ao construtor, o método 
     * isSuccessful retornar TRUE.
     * <p>
     * @return boolean importação da base feita com sucesso
     * @throws Exception 
     */
    public boolean importa() throws Exception {
        
        try {
            
            switch(type) {
                case HSSF_WORKBOOK:
                    importaXLS();
                    break;
                case XSSF_WORKBOOK:
                    importaXLSX();
                    break;
                case INVALID:
                    success = false;
                    errors.add("Tipo de arquivo inválido");
                    return success;
            }
            
            success = processaRows();
            
        } catch (Exception e) { 
            success = false;
            errors.add(e.getMessage());
            throw e;
        }
        return success;
    }

    protected void importaXLSX() throws Exception {
        
        try (
                Workbook workbook = StreamingReader.builder()
                    .rowCacheSize(rowCache)    // number of rows to keep in memory (defaults to 10)
                    .bufferSize(bufferSize)     // buffer size to use when reading InputStream to file (defaults to 1024)
                    .open(file);   
             ) {
            
            importaExcel(workbook);
            
        }
    }
    
    protected void importaXLS() throws Exception {
        
        try (
                FileInputStream fileIS = new FileInputStream(file);
            ) {
            
            Workbook workbook = new HSSFWorkbook(fileIS);
            
            importaExcel(workbook);
            
        }
        
    }
    
    private void importaExcel(Workbook workbook) throws Exception {
        
        if (singleSheet && workbook.getNumberOfSheets() > 1) 
            throw new Exception("O arquivo deve conter apenas uma Sheet.");

        if (firstSheetOnly) {
            Sheet sheet = workbook.getSheetAt(0);
            trataSheet(sheet, 1);
        } else {
            int sheetNo = 1;
            for (Sheet sheet : workbook) {
               trataSheet(sheet, sheetNo); 
               sheetNo++;
            }
        }
            
    }
    
    private void trataSheet(Sheet sheet, int sheetNo) {
        int lineNo = 0;
                
        try {
            for (Row r : sheet) {
                lineNo++;

                if (header && lineNo == 1) {
                    if (headerSize != r.getLastCellNum())
                        throw new Exception("A sheet #" + sheetNo + " não possui o número correto de colunas (" + headerSize + " colunas)!");
                    else {
                        continue;
                    }
                }

                trataRow(r, lineNo);

            }
        } catch (Exception e) {
            errors.add(e.getMessage());
        }
    }
    
    /**
     * Implementação do tratamento de uma linha do Excel.
     * <p>
     * Implementar método que popula os campos List e/ou Map para posterior 
     * iteração e escrita em banco, ou objetivo similar.
     * <p>
     * @param  r linha a ser processada.
     * @param lineNo número da linha sendo processada (utilizada para tratamento de erros).
     * @return boolean indicação de falha ou sucesso no processamento.
     */
    public abstract boolean trataRow(Row r, int lineNo);
    
    /**
     * Implementação do processamento dos objetos que representam as rows.
     * <p>
     * Implementar método que itera os campos List e/ou Map para escrita 
     * em banco, ou objetivo similar.
     * O boolean success deve ser modificado para false apenas em casos extremos,
     * como bloqueio de acesso ao banco.
     * O boolean success100percent deve ser mantido em true apenas se todas as 
     * linhas forem corretamente processadas.
     * <p>
     * @return boolean indicação de falha ou sucesso no processamento.
     */
    public abstract boolean  processaRows();
    
    private TYPE defineType(InputStream inp) throws IOException {
        try {
            InputStream inpBuf = FileMagic.prepareToCheckMagic(inp);
            if (FileMagic.valueOf(inpBuf).equals(FileMagic.OLE2)) {
                return TYPE.HSSF_WORKBOOK;
            } else if (FileMagic.valueOf(inpBuf).equals(FileMagic.OOXML)) {
                return TYPE.XSSF_WORKBOOK;
            }
            return TYPE.INVALID;
        } catch(Exception e) {
            e.printStackTrace();
            return TYPE.INVALID;
        }
    }
    
    /**
     * Método que indica se o upload do arquivo Excel foi feito corretamente
     * <p>
     * Chamar método importa apenas se este retornar TRUE.
     * <p>
     * @return boolean indica se upload foi feito corretamente.
     */
    public boolean isSuccessful() {
        return success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }

    public boolean isSuccess100percent() {
        return success100percent;
    }

    public void setSuccess100percent(boolean success100percent) {
        this.success100percent = success100percent;
    }

    public List<String> getErrors() {
        return errors;
    }

    public void setErrors(List errors) {
        this.errors = errors;
    }

    public TYPE getType() {
        return type;
    }

    public void setType(TYPE type) {
        this.type = type;
    }
    
    public void setSingleSheet() {
        this.singleSheet = true;
    }
    
    public void setMultipleSheet() {
        this.singleSheet = false;
    }
    
    public void readFirstSheetOnly() {
        this.firstSheetOnly = true;
    }
    
    public void readAllSheets() {
        this.firstSheetOnly = false;
    }
    
    public void setHasHeader() {
        this.header = true;
    }
    
    public void setNoHeader() {
        this.header = false;
    }
    
}
