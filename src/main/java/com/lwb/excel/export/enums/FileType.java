package com.lwb.excel.export.enums;

/**
 * 文件后缀
 * @author liuweibo
 * @date 2019/8/15
 */
public enum FileType {

    XLS("xls"),
    XLSX("xlsx");

    private String suffix;

    FileType(String suffix) {
        this.suffix = suffix;
    }

    public String getSuffix() {
        return suffix;
    }
}
