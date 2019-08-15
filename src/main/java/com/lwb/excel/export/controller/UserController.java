package com.lwb.excel.export.controller;

import com.lwb.excel.export.annotation.Export;
import com.lwb.excel.export.dao.UserMapper;
import com.lwb.excel.export.util.ExcelUtils;
import lombok.AccessLevel;
import lombok.experimental.FieldDefaults;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author liuweibo
 * @date 2019/8/14
 */
@FieldDefaults(level = AccessLevel.PRIVATE)
@RequestMapping("/user")
@RestController
public class UserController {

    @Autowired
    UserMapper userMapper;

    @GetMapping("/export/list")
    @Export("user-list.yml")
    public void exportList() {
        ExcelUtils.excel(() -> this.userMapper.getUserList());
    }
}
