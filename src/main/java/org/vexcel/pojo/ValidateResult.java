package org.vexcel.pojo;

public class ValidateResult {
    private Boolean success;
    private StringBuilder errorMsg ;

    public Boolean getSuccess() {
        return success;
    }

    public void setSuccess(Boolean success) {
        this.success = success;
    }

    public StringBuilder getErrorMsg() {
        return errorMsg;
    }

    public void setErrorMsg(StringBuilder errorMsg) {
        this.errorMsg = errorMsg;
    }
}
