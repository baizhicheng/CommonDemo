package com.bzc.example.demo.controller;

import com.bzc.example.demo.bean.ExportUserModel;
import com.bzc.example.demo.bean.UserModel;
import com.bzc.example.demo.constants.GlobalConstants;
import com.bzc.example.demo.service.UserModelService;
import com.bzc.example.demo.utils.ExportExcelUtil;
import com.bzc.example.demo.vo.ResultVo;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.UUID;

/**
 * @Author baizhicheng
 * @Date 2019/12/14 20:16
 */
@RestController
@Slf4j
@RequestMapping("/api/common")
public class UserModelController {

    @Resource
    private UserModelService service;

    //新增
    @RequestMapping(value = {"/insert"}, method = RequestMethod.POST)
    public ResultVo insert(HttpServletRequest request, HttpServletResponse response) {

        ResultVo retVo = new ResultVo();

        try {
            String username = request.getParameter("username");
            String age = request.getParameter("age");

            UserModel userModel = new UserModel();

            userModel.setId(UUID.randomUUID().toString());
            userModel.setUsername(username);
            userModel.setAge(age);

            service.insert(userModel);

            retVo.setMessage("新增成功！");
            retVo.setSuccess(true);
        } catch (Exception e) {
            retVo.setMessage("新增失败！");
            retVo.setSuccess(false);
            log.error(e.getMessage(), e);
        }

        return retVo;

    }

    //删除
    @RequestMapping(value = {"/delete"}, method = RequestMethod.POST)
    public ResultVo delete(HttpServletRequest request, HttpServletResponse response) {

        ResultVo retVo = new ResultVo();

        try {
            String id = request.getParameter("id");

            service.delete(id);

            retVo.setMessage("删除成功！");
            retVo.setSuccess(true);
        } catch (Exception e) {
            retVo.setMessage("删除失败！");
            retVo.setSuccess(false);
            log.error(e.getMessage(), e);
        }

        return retVo;

    }

    //修改
    @RequestMapping(value = {"/update"}, method = RequestMethod.POST)
    public ResultVo update(HttpServletRequest request, HttpServletResponse response) {

        ResultVo retVo = new ResultVo();

        try {
            String id = request.getParameter("id");
            String username = request.getParameter("username");
            String age = request.getParameter("age");

            UserModel userModel = new UserModel();

            userModel.setId(id);
            userModel.setUsername(username);
            userModel.setAge(age);

            service.update(userModel);

            retVo.setMessage("修改成功！");
            retVo.setSuccess(true);
        } catch (Exception e) {
            retVo.setMessage("修改失败！");
            retVo.setSuccess(false);
            log.error(e.getMessage(), e);
        }

        return retVo;

    }

    //分页查询
    @RequestMapping(value = {"/getPageList"}, method = RequestMethod.POST)
    public Map<String, Object> getPageList(HttpServletRequest request, HttpServletResponse response) {

        String pageNum = request.getParameter("pageNum");
        String pageSize = request.getParameter("pageSize");

        String username = request.getParameter("username");
        String age = request.getParameter("age");

        UserModel userModel = new UserModel();

        userModel.setUsername(username);
        userModel.setAge(age);

        Map<String, Object> jsonMap = service.getPageList(pageNum, pageSize,userModel);

        return jsonMap;

    }

    //根据ID获取对象
    @RequestMapping(value = {"/getUserModelById"}, method = RequestMethod.POST)
    public ResultVo getUserModelById(HttpServletRequest request, HttpServletResponse response) {

        ResultVo retVo = new ResultVo();

        try {
            String id = request.getParameter("id");

            UserModel userModel = service.getUserModelById(id);

            retVo.setData(userModel);
            retVo.setSuccess(true);
        } catch (Exception e) {
            retVo.setMessage("查询失败！");
            retVo.setSuccess(false);
            log.error(e.getMessage(), e);
        }

        return retVo;
    }

    //导出Excel表
    @RequestMapping(value = "/exportTable", method = RequestMethod.GET)
    public void exportBatteryTable(HttpServletRequest request, HttpServletResponse response) {

        String username = request.getParameter("username");
        String age = request.getParameter("age");

        UserModel userModel = new UserModel();

        userModel.setUsername(username);
        userModel.setAge(age);

        List<UserModel> userModelList = service.getList(userModel);

        List<ExportUserModel> exportTable = service.getExportList(userModelList);

        OutputStream os = null;
        try {
            response.addHeader("Content-Disposition","attachment;filename=" + System.currentTimeMillis() + ".xlsx");
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            os = response.getOutputStream();

            new ExportExcelUtil<ExportUserModel>().exportForExcel2007("用户表",
                    GlobalConstants.SHEET_TITLE, exportTable, os, GlobalConstants.DATE_FORMAT);

        } catch (IOException e) {
            log.error(e.getMessage(), e);
        } finally {
            try {
                if (null != os) os.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
    }

}
