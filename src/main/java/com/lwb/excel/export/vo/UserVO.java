package com.lwb.excel.export.vo;

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

/**
 * 用户vo
 * @author liuweibo
 * @date 2019/8/15
 */
@FieldDefaults(level = AccessLevel.PRIVATE)
@Data
public class UserVO {

    Long id;
    String name;
    ClassVO classVO;

}
