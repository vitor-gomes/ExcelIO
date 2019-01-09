package com.io.excel.utils;

import com.io.excel.annotations.ExcelColumn;
import java.lang.reflect.Field;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.Map;
import java.util.ResourceBundle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 *
 * @author pcoelho
 */
public class CellUtils {
    
    public static final DataFormatter DF = new DataFormatter();
    
    public static Date getCellDateValue(Cell cell, Field field, ResourceBundle bundle) throws Exception {
        String[] colPatterns = field.getAnnotation(ExcelColumn.class).columnDefinitions();
        if ( cell.getCellTypeEnum() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell))
            return cell.getDateCellValue();
        else {
            if (colPatterns == null || colPatterns.length == 0) {
                throw(new Exception("ColumnDefinition de um campo Date (" + com.io.excel.utils.StringUtils.getString(field.getAnnotation(ExcelColumn.class).name(), bundle) 
                    + ") de célula não numérica (" + field.getAnnotation(ExcelColumn.class).index() +   (cell.getRowIndex()+1) + ") não definido!"));
            } else {
                if (cell.getStringCellValue() == null || cell.getStringCellValue().equals("")) {
                    if (!field.getAnnotation(ExcelColumn.class).defaultValue().equals("")) {
                        for (String colPattern : colPatterns) {
                            try {
                                SimpleDateFormat dateFormatter = new SimpleDateFormat(colPattern);
                                return dateFormatter.parse(field.getAnnotation(ExcelColumn.class).defaultValue());
                            } catch(Exception e) {}
                        }
                    }
                    
                    if (field.getAnnotation(ExcelColumn.class).nullable())
                        return null;
                    
                } else {
                    
                    for (String colPattern : colPatterns) {
                        try {
                            SimpleDateFormat dateFormatter = new SimpleDateFormat(colPattern);
                            return dateFormatter.parse(cell.getStringCellValue());
                        } catch(Exception e) {}
                    }
                    
                }
                
            }
        }   
            throw(new Exception("Não foi possível parsear um campo Date (" + com.io.excel.utils.StringUtils.getString(field.getAnnotation(ExcelColumn.class).name(), bundle) + 
                    ") de uma célula não numérica (" + field.getAnnotation(ExcelColumn.class).index() +   (cell.getRowIndex()+1) + ") com os ColumnDefinitions passados!"));
    }
    
    public static Integer getCellIntegerValue(Cell cell, Field field, ResourceBundle bundle)  throws Exception {
        try {
            if (cell.getCellTypeEnum() == CellType.STRING) {
                if (cell.getStringCellValue() == null || cell.getStringCellValue().isEmpty()) {
                    if (field.getAnnotation(ExcelColumn.class).nullable())
                        return null;
                    else {
                        if (!field.getAnnotation(ExcelColumn.class).defaultValue().equals(""))
                            return NumberFormat.getInstance().parse(field.getAnnotation(ExcelColumn.class).defaultValue()).intValue();
                        else
                            throw(new Exception());
                    }
                } else {                    
                    return NumberFormat
                            .getInstance()
                            .parse(cell.getStringCellValue())
                            .intValue();
                }
            } else {
                return Integer.parseInt(DF.formatCellValue(cell).replaceAll("\\.", "").replaceAll(",", "").trim());
            }
        } catch (Exception e) {
            throw(new Exception("Não foi possível formatar a célula " + field.getAnnotation(ExcelColumn.class).index() +   (cell.getRowIndex()+1) + ", valor não é um inteiro!"));
        } 
    }
    
    public static int getCellIntValue(Cell cell, Field field, ResourceBundle bundle)  throws Exception {
        try {
            if (cell.getCellTypeEnum() == CellType.STRING) {
                if (cell.getStringCellValue() == null || cell.getStringCellValue().isEmpty()) {
                    if (!field.getAnnotation(ExcelColumn.class).defaultValue().equals(""))
                        return NumberFormat.getInstance().parse(field.getAnnotation(ExcelColumn.class).defaultValue()).intValue();
                    else
                        throw(new Exception());
                } else {                    
                    return NumberFormat
                            .getInstance()
                            .parse(cell.getStringCellValue())
                            .intValue();
                }
            } else {
                return Integer.parseInt(DF.formatCellValue(cell).replaceAll("\\.", "").replaceAll(",", "").trim());
            }
        } catch(Exception e) {
            throw(new Exception("Não foi possível formatar a célula " + field.getAnnotation(ExcelColumn.class).index() + (cell.getRowIndex()+1) + ", valor não é um inteiro!"));
        }
    }
    
    public static Double getCellDoubleValue(Cell cell, Field field, Locale locale, ResourceBundle bundle)  throws Exception {
        try {
            if (cell != null && cell.getCellTypeEnum() != CellType.BLANK && cell.getCellTypeEnum() == CellType.STRING) {
                if (cell.getStringCellValue() == null || cell.getStringCellValue().isEmpty()) {
                    if (field.getAnnotation(ExcelColumn.class).nullable())
                        return null;
                    else {
                        if (!field.getAnnotation(ExcelColumn.class).defaultValue().equals(""))
                            return NumberFormat.getInstance(locale).parse(field.getAnnotation(ExcelColumn.class).defaultValue()).doubleValue();
                        else
                            throw(new Exception());
                    }
                } else { 
                    if ((cell.getStringCellValue() == null || cell.getStringCellValue().equals("")) && !field.getAnnotation(ExcelColumn.class).nullable())
                        throw(new Exception());
                    
                    if (cell.getStringCellValue() == null || cell.getStringCellValue().equals(""))
                            return null;
                    
                    return NumberFormat
                            .getInstance(locale)
                            .parse(cell.getStringCellValue())
                            .doubleValue();
                }
            } else {
                if (cell == null || cell.getCellTypeEnum() == CellType.BLANK) {
                    if (!field.getAnnotation(ExcelColumn.class).nullable()) {
                        if (field.getAnnotation(ExcelColumn.class).defaultValue().equals(""))
                            throw(new Exception());
                        else
                            return NumberFormat.getInstance(locale).parse(field.getAnnotation(ExcelColumn.class).defaultValue()).doubleValue();
                    } else {
                        if (field.getAnnotation(ExcelColumn.class).defaultValue().equals(""))
                            return null;
                        else
                            return NumberFormat.getInstance(locale).parse(field.getAnnotation(ExcelColumn.class).defaultValue()).doubleValue();
                    }
                } else
                    return cell.getNumericCellValue();
            }
        } catch (Exception e) {
            throw(new Exception("Não foi possível formatar a célula " + field.getAnnotation(ExcelColumn.class).index() +   (cell.getRowIndex()+1) + ", valor não é numérico!"));
        } 
    }
    
    public static double getCellDoublePrimitiveValue(Cell cell, Field field, Locale locale, ResourceBundle bundle)  throws Exception {
        try {
            if (cell != null && cell.getCellTypeEnum() != CellType.BLANK && cell.getCellTypeEnum() == CellType.STRING) {
                if (cell.getStringCellValue() == null || cell.getStringCellValue().isEmpty()) {
                    if (!field.getAnnotation(ExcelColumn.class).defaultValue().equals(""))
                        return NumberFormat.getInstance(locale).parse(field.getAnnotation(ExcelColumn.class).defaultValue()).doubleValue();
                    else
                        throw(new Exception());
                } else {      
                    if ((cell.getStringCellValue() == null || cell.getStringCellValue().equals("")) && !field.getAnnotation(ExcelColumn.class).nullable())
                        throw(new Exception());
                    
                    return NumberFormat
                            .getInstance(locale)
                            .parse(cell.getStringCellValue())
                            .doubleValue();
                }
            } else {
                if (cell == null || cell.getCellTypeEnum() == CellType.BLANK) {
                    if (!field.getAnnotation(ExcelColumn.class).nullable() && field.getAnnotation(ExcelColumn.class).defaultValue().equals("")) {
                        throw(new Exception());
                    } else
                        return NumberFormat.getInstance(locale).parse(field.getAnnotation(ExcelColumn.class).defaultValue()).doubleValue();
                } else
                    return cell.getNumericCellValue();
            }
        } catch (Exception e) {
            throw(new Exception("Não foi possível formatar a célula " + field.getAnnotation(ExcelColumn.class).index() +   (cell.getRowIndex()+1) + ", valor não é numérico!"));
        } 
    }
    
    public static String getCellStringValue(Cell cell, Field field, ResourceBundle bundle)  throws Exception {
        try {
            if (cell == null) {
                return (field.getAnnotation(ExcelColumn.class).nullable() ? null : field.getAnnotation(ExcelColumn.class).defaultValue());
            }

            DataFormatter formatter = new DataFormatter();
            return formatter.formatCellValue(cell).trim();
        } catch (Exception e) {
            return "";
        }
    }

    public static Boolean getCellBooleanValue(Cell cell, Field field, Map<String, Boolean> booleanMap, ResourceBundle bundle)  throws Exception {
        try {
            if (cell.getCellTypeEnum() == CellType.BOOLEAN) 
                return cell.getBooleanCellValue();
            else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                Double doubleVal = cell.getNumericCellValue();
                switch(doubleVal.intValue()) {
                    case 0:
                        return false;
                    case 1:
                        return true;
                    default:
                        if (field.getAnnotation(ExcelColumn.class).nullable())
                            return null;
                        else
                            throw new Exception();
                }
            } else {
                DataFormatter formatter = new DataFormatter();
                String value =  formatter.formatCellValue(cell).trim();
                
                if (booleanMap != null) {
                    if(field.getAnnotation(ExcelColumn.class).nullable())
                        return booleanMap.get(value);
                    else if (booleanMap.get(value) == null)
                        throw new Exception();
                    else
                        return booleanMap.get(value);
                }
                
                return Boolean.parseBoolean(value);
            }
        } catch(Exception e) {
            throw new Exception("Não foi possível formatar a célula " + field.getAnnotation(ExcelColumn.class).index() +   (cell.getRowIndex()+1) + ", valor não é um booleano!");
        }
    }

    public static boolean getCellBooleanPrimitiveValue(Cell cell, Field field, Map<String, Boolean> booleanMap, ResourceBundle bundle)  throws Exception {
        try {
            if (cell.getCellTypeEnum() == CellType.BOOLEAN) 
                return cell.getBooleanCellValue();
            else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                Double doubleVal = cell.getNumericCellValue();
                switch(doubleVal.intValue()) {
                    case 0:
                        return false;
                    case 1:
                        return true;
                    default:
                        throw new Exception();
                }
            } else {
                DataFormatter formatter = new DataFormatter();
                String value =  formatter.formatCellValue(cell).trim();
                
                if (booleanMap != null) {
                    if (booleanMap.get(value) == null)
                        throw new Exception();
                    else
                        return booleanMap.get(value);
                }
                
                return Boolean.parseBoolean(value);
            }
        } catch(Exception e) {
            throw new Exception("Não foi possível formatar a célula " + field.getAnnotation(ExcelColumn.class).index() +   (cell.getRowIndex()+1) + ", valor não é um booleano!");
        }
    }
    
}
