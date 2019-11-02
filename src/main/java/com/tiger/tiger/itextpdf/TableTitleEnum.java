package com.tiger.tiger.itextpdf;

import java.util.Objects;

/**
 * Create by tiger on 2019/6/8
 */
public enum TableTitleEnum {

    COURSE_NAME(0, "课程名称"),
    COURSE_TYPE(1, "课程类型"),
    CREDIT(2, "学分"),
    SCORE(3, "成绩"),
    POINT(4, "绩点");

    private Integer code;
    private String msg;

    public Integer getCode() {
        return code;
    }

    public String getMsg() {
        return msg;
    }

    private TableTitleEnum(Integer code, String msg) {
        this.code = code;
        this.msg = msg;
    }

    public static TableTitleEnum getByCode(Integer code) {
        if (Objects.nonNull(code)) {
            for (TableTitleEnum value : TableTitleEnum.values()) {
                if (value.code == code) {
                    return value;
                }
            }
        }
        return null;
    }

}
