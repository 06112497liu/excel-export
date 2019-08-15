package com.lwb.excel.export.service.impl;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.lwb.excel.export.dao.UserMapper;
import com.lwb.excel.export.entity.User;
import com.lwb.excel.export.service.IUserService;
import org.springframework.stereotype.Service;

/**
 *  服务实现类
 * @author liuweibo
 * @date 2019/08/15
 */
@Service
public class UserServiceImpl extends ServiceImpl<UserMapper, User> implements IUserService {

}
