package com.lwb.excel.export.util;

import com.lwb.excel.export.exception.UtilException;
import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;

import java.util.List;
import java.util.Optional;

/**
 * @author liuweibo
 * @date 2019/8/14
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class ExcelConfig {

    /**
     * 导出文件名
     */
    String fileName;
    /**
     * 冻结规则
     * </p>
     * 例如：0,2,1,2表示冻结两行，首行可见序号为2
     * 第一个数：冻结几列
     * 第二个数：冻结几行
     * 第三个数：首列可见序号，从1开始
     * 第三个数：首行可见序号，从1开始
     */
    String freezePaneIndex;
    /**
     * 表头信息
     * </p>
     * 支持多行表头
     */
    List<List<Header>> headers;

    List<String> fields;

    @Data
    @FieldDefaults(level = AccessLevel.PRIVATE)
    static class Header {
        /**
         * 表头名称
         */
        String name;
        /**
         * 单元格合并规则
         * </p>
         * 例如：1,2,3,4表示合并单元格第一行和第二行的第三列和第四列
         */
        String mergeIndex;
    }

    /**
     * 校验配置的完整性
     */
    public void validate() {
        Optional.ofNullable(this)
            .filter(config -> CollectionUtils.isNotEmpty(this.getHeaders()))
            .filter(config -> StringUtils.isNotEmpty(this.fileName))
            .orElseThrow(() -> new UtilException("导出excel配置信息不完整"));
    }

}
