package com.bzc.example.demo.controller;

import com.bzc.example.demo.bean.ExportUserModel;
import com.bzc.example.demo.bean.UserModel;
import com.bzc.example.demo.constants.GlobalConstants;
import com.bzc.example.demo.service.UserModelService;
import com.bzc.example.demo.utils.ExcelUtils;
import com.bzc.example.demo.utils.ExportExcelUtil;
import com.bzc.example.demo.utils.FreeMarkerUtils;
import com.bzc.example.demo.utils.StringUtils;
import com.bzc.example.demo.vo.ResultVo;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.annotation.Resource;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

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

    @Value("${downloadpath}")
    String path;

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
                    GlobalConstants.SHEET_EXPORT_TITLE, exportTable, os, GlobalConstants.DATE_FORMAT);

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

    //导入Excel表
    @RequestMapping(value = {"/importTable"}, method = RequestMethod.POST)
    public ResultVo importFile(HttpServletRequest request, HttpServletResponse response) {
        ResultVo resultVo;
        Map<String, Object> map = new HashMap<>();
        InputStream fileInput = null;

        try {
            MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;
            MultipartFile file = multipartRequest.getFile("file");// 获取上传文件
            fileInput = file.getInputStream();
            String filename = file.getOriginalFilename();
            Workbook workbook = WorkbookFactory.create(fileInput);

            map = this.service.importTable(workbook, filename, request);

            if (null == map)
                return ResultVo.error("", "导入错误，请重新导入数据！");
            else
                map.put(ExcelUtils.SAVE_FILE, "");

            // 导入模板校验
            boolean errorTitle = (boolean) map.get(ExcelUtils.SHEET_TITLE_ERROR);

            if (errorTitle)
                return ResultVo.error("", "模板错误，请使用正确的导入模板！");

            int totalCount = (int) map.get(ExcelUtils.TOTAL_COUNT);

            if (0 == totalCount)
                resultVo = ResultVo.error("", "请填写需要导入的数据！");
            else {
                resultVo = ResultVo.success(map);
                resultVo.setMessage("数据导入成功！");
            }

        } catch (Exception e) {
            log.error(e.getMessage(), e);
            resultVo = ResultVo.error(GlobalConstants.SERVER_ERROR_CODE, "数据导入失败！");
        }finally {
            if (null != fileInput) {
                try {
                    fileInput.close();
                } catch (IOException e) {
                    log.error(e.getMessage(), e);
                }
            }
        }

        return resultVo;
    }

    // 导出为Word
    @RequestMapping(value = {"/exportWord"}, method = RequestMethod.GET)
    public void exportWord(HttpServletRequest request,HttpServletResponse response) throws ParseException{
        String username = request.getParameter("username");
        String age = request.getParameter("age");

        UserModel userModel = new UserModel();

        userModel.setUsername(username);
        userModel.setAge(age);

        SimpleDateFormat df = new SimpleDateFormat("yyyyMMddHHmmss");//设置日期格式
        String time = df.format(new Date());

        String filename="UserModel"+time+".doc";

        FreeMarkerUtils<UserModel> userModelFreeMarkerUtils = new FreeMarkerUtils<>(new UserModel());
        try {
            //集合用parseCollection，单个字串或对象用parseSingleBean
            Map<String, Object> dataMap = userModelFreeMarkerUtils.parseCollection("userModelData", service.getList(userModel));

            FreeMarkerUtils.generateReports(dataMap, "exportWordTemplate.ftl", filename);
            download(filename,response);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
    }

    private void download(String filename,HttpServletResponse response) {
        try {
            String filepath=path+filename;
            // path是指欲下载的文件的路径
            File file = new File(filepath);
            // 取得文件名
            filename = java.net.URLEncoder.encode(filename,"utf-8");
            // 清空response
            response.reset();
            // 设置response的Header
            response.addHeader("Content-Disposition", "attachment;filename="
                    + new String(filename.getBytes()));
            response.addHeader("Content-Length", "" + file.length());
            response.setCharacterEncoding("utf-8");
            OutputStream toClient = new BufferedOutputStream(
                    response.getOutputStream());
            response.setContentType("application/msword;charset=utf-8");
            // 以流的形式下载文件
            InputStream fis = new BufferedInputStream(new FileInputStream(filepath));
            byte[] buffer = new byte[fis.available()];
            fis.read(buffer);
            fis.close();
            toClient.write(buffer);
            toClient.flush();
            toClient.close();
        } catch (IOException ex) {
            log.error(ex.getMessage(),ex);
        }

    }

}
