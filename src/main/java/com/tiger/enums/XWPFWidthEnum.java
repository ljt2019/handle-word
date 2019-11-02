package com.tiger.enums;

import java.util.Objects;

/**
 * XWPF 表格列宽比例
 *
 * @author: tiger
 * @create: 2019-11-02 12:20
 */
public enum XWPFWidthEnum {
    NAME(0, "2400"),
    TYPE(1, "1200"),
    CREDIT(2, "500"),
    SCORE(3, "500"),
    GPA(4, "500"),

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
