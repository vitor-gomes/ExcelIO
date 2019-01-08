package com.io.excel.utils;

import java.util.MissingResourceException;
import java.util.ResourceBundle;

/**
 *
 * @author pcoelho
 */
public class StringUtils {
    
    static public String getString( String key, ResourceBundle bundle )
    {
        try
        {
            return( bundle.getString( key ) );
        }
        catch( Exception ex )
        {
            return( key );
        }
    }
    
}
