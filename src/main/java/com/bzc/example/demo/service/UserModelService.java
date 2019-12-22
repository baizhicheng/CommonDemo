package com.bzc.example.demo.service;

import com.bzc.example.demo.bean.ExportUserModel;
import com.bzc.example.demo.bean.UserModel;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.http.HttpServletRequest;
import java.util.List;
import java.util.Map;

/**
 * @Author baizhicheng
 * @Date 2019/12/14 19:39
 */
public interface UserModelService {

    //新增
    void insert(UserModel user);

    //删除
    void delete(String id);

    //修改
    void update(UserModel user);

    //查询所有记录(Map)
    Map<String, Object> getPageList(String pageNum, String pageSize, UserModel user);

    //根据ID获取对象
    UserModel getUserModelById(String id);

    //查询所有记录(List)
    List<UserModel> getList(UserModel user);

    //导出数据
    List<ExportUserModel> getExportList(List<UserModel> list);

    //导入数据
    Map<String, Object> importTable(Workbook workbook, String filename, HttpServletRequest request);

}
