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

    //查询所有记录(有参数)
    List<UserModel> getList(UserModel user);

    //查询所有记录(无参数)
    List<UserModel> getList();

    //根据ID获取对象
    UserModel getUserModelById(String id);

    //从Excel录入数据
    void insertByExcel(List<UserModel> list);

    //从Excel更新数据
    void updateByExcel(List<UserModel> list);

}
