package org.vexcel.cache;

import org.vexcel.exception.ValidateRuntimeException;
import org.vexcel.pojo.ExcelConfig;
import org.vexcel.tools.XmlUtils;

import java.util.HashMap;

public class ConfigCache {
    private static HashMap<String, ExcelConfig> configCache = null;
    public  static synchronized HashMap<String, ExcelConfig>  getExcelConfig(){
        if(configCache==null){
            configCache =   new XmlUtils().getAllValidators();
            return configCache;
        }else{
            return configCache;
        }
    }

    public  static  ExcelConfig  getExcelConfig(String validatorId){
        HashMap<String, ExcelConfig>  configs =  getExcelConfig();
        if(configs.containsKey(validatorId)){
            return configs.get(validatorId);
        }else{
            throw new ValidateRuntimeException("根据id未找到配置");
        }
    }

}
