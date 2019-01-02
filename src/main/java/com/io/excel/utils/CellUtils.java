package com.io.excel.utils;

import com.io.excel.annotations.ExcelColumn;
import java.lang.reflect.Field;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 *
 * @author pcoelho
 */
public class CellUtils {
    
    // TODO configurable Locale
    public static final Locale LOCALE = new Locale("pt", "BR");
    public static final DataFormatter DF = new DataFormatter();
    
    public static Date getCellDateValue(Cell cell, Field field) throws Exception {
        String[] colPatterns = field.getAnnotation(ExcelColumn.class).columnDefinitions();
        String colString = field.getAnnotation(ExcelColumn.class).index();
        if ( cell.getCellTypeEnum() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell))
            return cell.getDateCellValue();
        else {
            if (colPatterns == null || colPatterns.length == 0)
                throw(new Exception("ColumnDefinition de um campo Date (" + field.getName() + ") de coluna não numérica (" + colString + ") não definido!"));

            for (String colPattern : colPatterns) {
                try {
                    SimpleDateFormat dateFormatter = new SimpleDateFormat(colPattern);
                    return dateFormatter.parse(cell.getStringCellValue());
                } catch(Exception e) {}
            }
            
            throw(new Exception("Não foi possível parsear um campo Date (" + field.getName() + ") de uma coluna não numérica (" + colString + ") com os ColumnDefinitions passados!"));
        }
    }
    
    public static int getCellIntValue(Cell cell, Field field)  throws Exception {
        try {
            return Integer.parseInt(DF.formatCellValue(cell).replaceAll(".", "").replaceAll(",", "").trim());
        } catch(Exception e) {
            throw(new Exception("Não foi possível formatar a célula " + field.getAnnotation(ExcelColumn.class).index() + cell.getAddress().getRow() + ", valor não é um inteiro!"));
        }
    }
    
    public static double getCellDoubleValue(Cell cell, Field field)  throws Exception {
        try {
            if (cell.getCellTypeEnum() == CellType.STRING) {
                if (cell.getStringCellValue() == null || cell.getStringCellValue().isEmpty()) {
                    return 0;
                } else {                    
                    return NumberFormat
                            .getInstance(LOCALE)
                            .parse(cell.getStringCellValue())
                            .doubleValue();
                }
            } else {
                return cell.getNumericCellValue();
            }
        } catch (Exception e) {
            throw(new Exception("Não foi possível formatar a célula " + field.getAnnotation(ExcelColumn.class).index() + cell.getAddress().getRow() + ", valor não é numérico!"));
        } 
    }
    
    public static String getCellStringValue(Cell cell, Field field)  throws Exception {
        try {
            if (cell == null) {
                return null;
            }

            DataFormatter formatter = new DataFormatter();
            return formatter.formatCellValue(cell).trim();
        } catch (Exception e) {
            return "";
        }
    }
    
}
