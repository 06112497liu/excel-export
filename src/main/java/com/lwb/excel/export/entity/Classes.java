package com.lwb.excel.export.entity;

import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

import java.io.Serializable;

/**
 * 实体
 * @author liuweibo
 * @date 2019/08/15
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@TableName("b_classes")
public class Classes implements Serializable {

    private static final long serialVersionUID = 1L;
    private Long id;
    private String name;
    private Long schoolId;


}
