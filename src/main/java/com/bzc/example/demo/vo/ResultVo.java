package com.bzc.example.demo.vo;

import com.bzc.example.demo.bean.Error;
import java.io.Serializable;
import java.util.Arrays;
import java.util.List;

/** @deprecated */
@Deprecated
public class ResultVo<T> implements Serializable {
    private static final long serialVersionUID = 4050671089513293158L;
    /** @deprecated */
    @Deprecated
    private String errorCode;
    /** @deprecated */
    @Deprecated
    private String message;
    private T data;
    private boolean success;
    private List<Error> errors;

    public ResultVo() {
        this.success = true;
        this.errors = Arrays.asList();
    }

    public ResultVo(T data) {
        this(data, (String)null, (String)null, true);
    }

    /** @deprecated */
    @Deprecated
    public ResultVo(String errorCode, String message) {
        this((T) null, errorCode, message, false);
    }

    public ResultVo(T data, String errorCode, String message, boolean success) {
        this.success = true;
        this.errors = Arrays.asList();
        this.data = data;
        this.errorCode = errorCode;
        this.message = message;
        this.success = success;
    }

    /** @deprecated */
    @Deprecated
    public ResultVo<T> message(String message) {
        this.setMessage(message);
        return this;
    }

    public static <T> ResultVo<T> success(T data) {
        return new ResultVo(data);
    }

    public static <T> ResultVo<T> error(Error... errors) {
        ResultVo<T> resultVo = new ResultVo();
        resultVo.setSuccess(false);
        if (errors != null && errors.length > 0) {
            resultVo.setErrors(Arrays.asList(errors));
        }

        return resultVo;
    }

    /** @deprecated */
    @Deprecated
    public static <T> ResultVo<T> error(String errorCode, String message) {
        ResultVo<T> resultVo = new ResultVo(errorCode, message);
        resultVo.setErrors(Arrays.asList(new Error(errorCode, message)));
        return resultVo;
    }

    /** @deprecated */
    @Deprecated
    public static <T> ResultVo<T> error(String errorCode, Error... errors) {
        ResultVo<T> resultVo = new ResultVo(errorCode, (String)null);
        if (errors != null && errors.length > 0) {
            resultVo.setErrors(Arrays.asList(errors));
        }

        return resultVo;
    }

    /** @deprecated */
    @Deprecated
    public static <T> ResultVo<T> error(String errorCode, List<Error> errors) {
        ResultVo<T> resultVo = new ResultVo(errorCode, (String)null);
        resultVo.setErrors(errors);
        return resultVo;
    }

    /** @deprecated */
    @Deprecated
    public String getErrorCode() {
        return this.errorCode;
    }

    /** @deprecated */
    @Deprecated
    public void setErrorCode(String errorCode) {
        this.errorCode = errorCode;
    }

    /** @deprecated */
    @Deprecated
    public String getMessage() {
        return this.message;
    }

    /** @deprecated */
    @Deprecated
    public void setMessage(String message) {
        this.message = message;
    }

    public T getData() {
        return this.data;
    }

    public void setData(T data) {
        this.data = data;
    }

    public boolean isSuccess() {
        return this.success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }

    public List<Error> getErrors() {
        return this.errors;
    }

    public void setErrors(List<Error> errors) {
        this.errors = errors;
    }
}

