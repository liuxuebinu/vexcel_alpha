package org.vexcel.tools;

import org.vexcel.pojo.VSheet;
import org.vexcel.pojo.ValidateResult;

import java.io.InputStream;
import java.util.List;

public class ExcelValidUtil {
   public static ValidateResult doValidate(InputStream ios , String validatorId){
       List<VSheet> rules = new  XmlUtils().getRuleByName(validatorId);
       String xmlType = new  XmlUtils().getTypeById(validatorId);
       return ExcelUtils.readExcel(ios,rules,xmlType) ;
   }
}
