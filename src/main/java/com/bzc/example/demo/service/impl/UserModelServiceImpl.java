package com.bzc.example.demo.service.impl;

import com.bzc.example.demo.bean.UserModel;
import com.bzc.example.demo.dao.UserModelMapper;
import com.bzc.example.demo.service.UserModelService;
import com.bzc.example.demo.utils.StringUtils;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author baizhicheng
 * @Date 2019/12/14 19:39
 */
@Transactional
@Service
@Slf4j
public class UserModelServiceImpl implements UserModelService {

    @Resource
    private UserModelMapper dao;

    //新增
    @Override
    public void insert(UserModel user){
        dao.insert(user);
    }

    //删除
    @Override
    public void delete(String id){
        dao.delete(id);
    }

    //修改
    @Override
    public void update(UserModel user){
        dao.update(user);
    }

    //分页查询
    @Override
    public Map<String, Object> getPageList(String pageNum, String pageSize, UserModel user){
        int pagenum = 1;
        int pagesize = 10;

        if (StringUtils.isInteger(pageNum)){
            pagenum = Integer.parseInt(pageNum);
        }

        if (StringUtils.isInteger(pageSize)){
            pagesize = Integer.parseInt(pageSize);
        }

        //PageHelper分页插件使用
        PageHelper.startPage(pagenum, pagesize);
        List<UserModel> userModelList = dao.getList(user);

        //将数据库查出的值扔到PageInfo里实现分页效果
        PageInfo<UserModel> pageInfo = new PageInfo<>(userModelList);

        //将结果展示到map里
        Map<String, Object> jsonMap = new HashMap<String, Object>();

        jsonMap.put("ret", "0");
        jsonMap.put("msg", "SUCCESS");
        jsonMap.put("body", userModelList);//数据结果
        jsonMap.put("total", pageInfo.getTotal());//获取数据总数
        jsonMap.put("pageSize", pageInfo.getPageSize());//获取长度
        jsonMap.put("pageNum", pageInfo.getPageNum());//获取当前页数

        return jsonMap;
    }

    //根据ID获取对象
    @Override
    public UserModel getUserModelById(String id){
        return dao.getUserModelById(id);
    }

}
