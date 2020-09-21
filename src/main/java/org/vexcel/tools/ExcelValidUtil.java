package org.vexcel.tools;

import org.vexcel.cache.ConfigCache;
import org.vexcel.pojo.ExcelConfig;
import org.vexcel.pojo.VSheet;
import org.vexcel.pojo.ValidateResult;

import java.io.InputStream;
import java.util.List;

public class ExcelValidUtil {
   public static ValidateResult doValidate(InputStream ios , String validatorId){
       ExcelConfig config = ConfigCache.getExcelConfig(validatorId);
       List<VSheet> rules = config.getSheets();
       String xmlType = config.getExcelType();
       return ExcelUtils.readExcel(ios,rules,xmlType) ;
   }
}
