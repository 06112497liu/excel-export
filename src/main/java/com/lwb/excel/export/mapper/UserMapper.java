package com.lwb.excel.export.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import com.lwb.excel.export.entity.User;
import com.lwb.excel.export.vo.UserVO;

import java.util.List;

/**
 *  Mapper 接口
 * @author liuweibo
 * @date 2019/08/15
 */
public interface UserMapper extends BaseMapper<User> {

    List<UserVO> getUserList();

}
