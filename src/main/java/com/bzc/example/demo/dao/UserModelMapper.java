package com.bzc.example.demo.dao;

import com.bzc.example.demo.bean.UserModel;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

/**
 * @Author baizhicheng
 * @Date 2019/12/14 19:00
 */

@Mapper
public interface UserModelMapper {

    //新增
    void insert(UserModel user);

    //删除
    void delete(String id);

    //修改
    void update(UserModel user);

    //查询所有记录
    List<UserModel> getList(UserModel user);

    //根据ID获取对象
    UserModel getUserModelById(String id);

}
