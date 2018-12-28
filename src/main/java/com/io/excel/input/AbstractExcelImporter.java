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
    protected static String outputDirPath;
    
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
    
    
    /**
     * TODO
     * @param path 
     */
    public AbstractExcelImporter(String path) {
        this(path, 0);
    }
    
    /**
     * TODO
     * @param path
     * @param headersize 
     */
    public AbstractExcelImporter(String path, int headersize) {
        this(null, null, F_ROWCACHE, F_BUFFERSIZE, headersize);
        this.file = new File(path);
    }
    
    /**
     * TODO
     * @param path
     * @param bufferSize
     * @param rowCache 
     */
    public AbstractExcelImporter(String path, int bufferSize, int rowCache) {
        this(path, rowCache, bufferSize, 0);
    }
    
    /**
     * TODO
     * @param path
     * @param rowCache
     * @param bufferSize
     * @param headersize 
     */
    public AbstractExcelImporter(String path, int rowCache, int bufferSize, int headersize) {
        this(null, null, rowCache, bufferSize, headersize);
        this.file = new File(path);
    }
    
    /**
     * TODO
     * @param fileObject
     * @param outputDirPath 
     */
    public AbstractExcelImporter(Object fileObject, String outputDirPath) {
        this(fileObject, outputDirPath, 0);
    }
    
    /**
     * TODO
     * @param fileObject
     * @param outputDirPath
     * @param headersize 
     */
    public AbstractExcelImporter(Object fileObject, String outputDirPath, int headersize) {
        this(fileObject, outputDirPath, F_ROWCACHE, F_BUFFERSIZE, headersize);
    }
    
    /**
     * TODO
     * @param fileObject
     * @param outputDirPath
     * @param rowCache
     * @param bufferSize 
     */
    public AbstractExcelImporter(Object fileObject, String outputDirPath, int rowCache, int bufferSize) {
        this(fileObject, outputDirPath, rowCache, bufferSize,0);
    }
    
    /**
     * TODO
     * @param fileObject
     * @param outputDirPath
     * @param rowCache
     * @param bufferSize
     * @param headersize 
     */
    public AbstractExcelImporter(Object fileObject, String outputDirPath, int rowCache, int bufferSize, int headersize) {
        this.rowCache = rowCache;
        this.bufferSize = bufferSize;
        this.headerSize = headersize;
        this.fileObject = fileObject;
        this.outputDirPath = outputDirPath;
        try {
            if (fileObject != null & outputDirPath != null) {
                upload();
                success = true;
            }
        } catch (Exception e) {
            e.printStackTrace();
            success = false;
            errors.add(e.getMessage());
        }
    }
    
    private void upload() throws Exception {
        fileItem = (FileItem) fileObject;
        String fileName = fileItem.getName();
        
        fileName = fileName.contains("\\") ? fileName.substring(fileName.lastIndexOf("\\") + 1) : fileName;
        
        if (fileName.length() > 0) {
            try{
                String path = outputDirPath + File.separator + fileName;

                uploadedStream = fileItem.getInputStream();

                type = defineType(uploadedStream);

                switch(type) {
                    case INVALID:
                        //TODO: set errors string elsewhere!!!!
                        throw(new Exception("Invalid file!"));
                    case HSSF_WORKBOOK:
                        writeXLS(path);
                    case XSSF_WORKBOOK:
                        writeXLSX(path);
                        break;
                }
            
            } catch (Exception e) { throw e; }
            finally { uploadedStream.close(); }
            
        } else {
            throw(new Exception("Arquivo Inexistente!"));
        }
        
    }
    
    /**
     * TODO
     * Método de escrita do arquivo em disco.
     * <p>
     * Este método escreve o arquivo XLS em disco.
     * Dar override neste método da implementação da classe quando desejar bloquear
     * o upload de arquivos XLS
     * <p>
     * @param path caminho completo do arquivo a ser escrito
     * @throws Exception 
     */
    protected void writeXLS(String path) throws Exception {
        write(path);
    }
    
    /**
     * TODO
     * Método de escrita do arquivo em disco.
     * <p>
     * Este método escreve o arquivo XLSX em disco.
     * Dar override neste método da implementação da classe quando desejar bloquear
     * o upload de arquivos XLSX
     * <p>
     * @param path caminho completo do arquivo a ser escrito
     * @throws Exception 
     */
    protected void writeXLSX(String path) throws Exception {
        write(path);
    }
    
    private void write(String path) throws Exception {
        
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
     * TODO
     * Método de importação dos dados do Excel
     * <p>
     * Método que deve ser chamado caso após a chamada ao construtor, o método 
     * isSuccessful retornar TRUE.
     * <p>
     * @return boolean importação da base feita com sucesso
     * @throws Exception 
     */
    public boolean importFile() throws Exception {
        
        try {
            
            if (type == null) {
                InputStream inputStream = new FileInputStream(file);
                type = defineType(inputStream);
            }
            
            switch(type) {
                case HSSF_WORKBOOK:
                    importXLS();
                    break;
                case XSSF_WORKBOOK:
                    importXLSX();
                    break;
                case INVALID:
                    success = false;
                    errors.add("Tipo de arquivo inválido");
                    return success;
            }
            
            success = handleRows();
            
        } catch (Exception e) { 
            success = false;
            errors.add(e.getMessage());
            throw e;
        }
        return success;
    }

    protected void importXLSX() throws Exception {
        
        try (
                Workbook workbook = StreamingReader.builder()
                    .rowCacheSize(rowCache)    // number of rows to keep in memory (defaults to 10)
                    .bufferSize(bufferSize)     // buffer size to use when reading InputStream to file (defaults to 1024)
                    .open(file);   
             ) {
            
            importExcel(workbook);
            
        }
    }
    
    protected void importXLS() throws Exception {
        
        try (
                FileInputStream fileIS = new FileInputStream(file);
            ) {
            
            Workbook workbook = new HSSFWorkbook(fileIS);
            
            importExcel(workbook);
            
        }
        
    }
    
    private void importExcel(Workbook workbook) throws Exception {
        
        if (singleSheet && workbook.getNumberOfSheets() > 1) 
            throw new Exception("O arquivo deve conter apenas uma Sheet.");

        if (firstSheetOnly) {
            Sheet sheet = workbook.getSheetAt(0);
            handleSheet(sheet, 1);
        } else {
            int sheetNo = 1;
            for (Sheet sheet : workbook) {
               handleSheet(sheet, sheetNo); 
               sheetNo++;
            }
        }
            
    }
    
    private void handleSheet(Sheet sheet, int sheetNo) {
        int lineNo = 0;
                
        try {
            for (Row r : sheet) {
                lineNo++;

                if (header && lineNo == 1) {
                    if (headerSize != 0 && headerSize != r.getLastCellNum())
                        throw new Exception("A sheet #" + sheetNo + " não possui o número correto de colunas (" + headerSize + " colunas)!");
                    else {
                        continue;
                    }
                }

                handleRow(r, lineNo);

            }
        } catch (Exception e) {
            errors.add(e.getMessage());
        }
    }
    
    /**
     * TODO
     * Implementação do tratamento de uma linha do Excel.
     * <p>
     * Implementar método que popula os campos List e/ou Map para posterior 
     * iteração e escrita em banco, ou objetivo similar.
     * <p>
     * @param  r linha a ser processada.
     * @param lineNo número da linha sendo processada (utilizada para tratamento de erros).
     * @return boolean indicação de falha ou sucesso no processamento.
     */
    public abstract boolean handleRow(Row r, int lineNo);
    
    /**
     * TODO
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
    public abstract boolean  handleRows();
    
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
