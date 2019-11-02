package com.tiger.enums;

import java.util.Objects;

/**
 * @description:
 * @author: tiger
 * @create: 2019-11-02 12:20
 */
public enum TitleEnum {
    NAME(0, "课程名称"),
    TYPE(1, "课程类型"),
    CREDIT(2, "学分"),
    SCORE(3, "成绩"),
    GPA(4, "绩点"),

    ;

    private Integer code;
    private String  msg;

    public Integer getCode() {
        return code;
    }

    public String getMsg() {
        return msg;
    }

    private TitleEnum(Integer code, String msg) {
        this.code = code;
        this.msg = msg;
    }

    public static TitleEnum getByCode(Integer code) {
        if (Objects.nonNull(code)) {
            for (TitleEnum value : TitleEnum.values()) {
                if (value.code.equals(code)) {
                    return value;
                }
            }
        }
        return null;
    }
}
