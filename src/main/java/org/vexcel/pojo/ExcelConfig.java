package org.vexcel.pojo;

import java.util.List;

public class ExcelConfig {
    private String xmlPath;
    private String excelType;
    private String validatorId;
    private List<VSheet> sheets ;

    public String getXmlPath() {
        return xmlPath;
    }

    public void setXmlPath(String xmlPath) {
        this.xmlPath = xmlPath;
    }

    public String getExcelType() {
        return excelType;
    }

    public void setExcelType(String excelType) {
        this.excelType = excelType;
    }

    public String getValidatorId() {
        return validatorId;
    }

    public void setValidatorId(String validatorId) {
        this.validatorId = validatorId;
    }

    public List<VSheet> getSheets() {
        return sheets;
    }

    public void setSheets(List<VSheet> sheets) {
        this.sheets = sheets;
    }
}
