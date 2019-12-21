package com.bzc.example.demo.service;

import com.bzc.example.demo.bean.UserModel;

import java.util.List;
import java.util.Map;

/**
 * @Author baizhicheng
 * @Date 2019/12/14 19:39
 */
public interface UserModelService {

    //新增
    public void insert(UserModel user);

    //删除
    public void delete(String id);

    //修改
    public void update(UserModel user);

    //查询所有记录
    public Map<String, Object> getPageList(String pageNum, String pageSize, UserModel user);

    //根据ID获取对象
    public UserModel getUserModelById(String id);

}
