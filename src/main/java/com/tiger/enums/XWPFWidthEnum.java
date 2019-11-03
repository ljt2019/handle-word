package com.tiger.enums;

import java.util.Objects;

/**
 * XWPF 表格列宽比例
 *
 * @author: tiger
 * @create: 2019-11-02 12:20
 */
public enum XWPFWidthEnum {

    //总宽度要大于纸张宽度，比例效果才好
    NAME(0, "24000"),
    TYPE(1, "12000"),
    CREDIT(2, "5000"),
    SCORE(3, "5000"),
    GPA(4, "5000"),

    ;

    private Integer code;
    private String  msg;

    public Integer getCode() {
        return code;
    }

    public String getMsg() {
        return msg;
    }

    private XWPFWidthEnum(Integer code, String msg) {
        this.code = code;
        this.msg = msg;
    }

    public static XWPFWidthEnum getByCode(Integer code) {
        if (Objects.nonNull(code)) {
            for (XWPFWidthEnum value : XWPFWidthEnum.values()) {
                if (value.code.equals(code)) {
                    return value;
                }
            }
        }
        return null;
    }
}
