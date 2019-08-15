package com.lwb.excel.export.vo;

import com.lwb.excel.export.entity.School;
import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

/**
 * @author liuweibo
 * @date 2019/8/15
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class ClassVO {

    Long id;
    String name;
    School school;
}
