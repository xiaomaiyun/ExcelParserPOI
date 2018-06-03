package com.chenjian.utils;

public class WDWUtil {

    /**
     * @描述：判断是否是Excel2003，返回true表示是Excel2003
     * @param filePath
     * @return
     */
    public static boolean isExcel2003(String filePath){
        return filePath.matches("^.+\\.(?i)(xls)$");
    }

    /**
     * @描述：判断是否是Excel2007，返回true表示是Excel2007
     * @param filePath
     * @return
     */
    public static boolean isExcel2007(String filePath){
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }
    /**
     * @描述：验证是否是EXCEL文件
     * @param filePath
     * @return
     */
    public static boolean validateExcel(String filePath){
        if ((isExcel2003(filePath))||(isExcel2007(filePath))){
            return true;
        }
        return false;
    }
}
