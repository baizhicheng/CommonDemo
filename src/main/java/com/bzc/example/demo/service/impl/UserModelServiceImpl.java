package com.bzc.example.demo.service.impl;

import com.bzc.example.demo.bean.ExportUserModel;
import com.bzc.example.demo.bean.ImportUserModel;
import com.bzc.example.demo.bean.UserModel;
import com.bzc.example.demo.constants.GlobalConstants;
import com.bzc.example.demo.dao.UserModelMapper;
import com.bzc.example.demo.service.UserModelService;
import com.bzc.example.demo.utils.ExcelUtils;
import com.bzc.example.demo.utils.StringUtils;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

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

    @Resource
    private ExcelUtils excelUtils;

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

    //查询所有记录(List)
    @Override
    public List<UserModel> getList(UserModel user){
        return dao.getList(user);
    }

    //导出数据
    @Override
    public List<ExportUserModel> getExportList(List<UserModel> list){
        List<ExportUserModel> exports = new ArrayList<>();

        for (UserModel userModel : list) {
            ExportUserModel export = new ExportUserModel();

            export.setId(userModel.getId());
            export.setUsername(userModel.getUsername());
            export.setAge(userModel.getAge());

            exports.add(export);
        }

        return exports;
    }

    //导入数据
    public Map<String, Object> importTable(Workbook workbook, String filename, HttpServletRequest request){
        Map<String, Object> map = null;

        try{
            map = excelUtils.parseExcel(workbook, filename, ImportUserModel.class);

            if (null == map){
                return null;
            }

            String[] titles = (String[]) map.get(ExcelUtils.SHEET_TITLES);

            // 验证标题
            if (!Arrays.equals(GlobalConstants.SHEET_IMPORT_TITLE, titles)) {
                map.put(ExcelUtils.SHEET_TITLE_ERROR, true);
                map.put(ExcelUtils.IS_SUCCESS, false);
                map.put(ExcelUtils.MSG, "标题错误，请参考模板！");

                return map;
            }

            map.put(ExcelUtils.SHEET_TITLE_ERROR, false);

            if (!map.containsKey(ExcelUtils.VALID_LIST)){
                return map;
            }

            List<ImportUserModel> list = (List<ImportUserModel>) map.get(ExcelUtils.VALID_LIST);

            if (list.size() == 0){
                return map;
            }

            List<UserModel> userInsertList = new ArrayList<>();

            List<UserModel> userUpdateList = new ArrayList<>();

            for(int i = 0 ;i< list.size() ; i++){
                UserModel userModel = new UserModel();

                userModel.setId(list.get(i).getId());
                userModel.setUsername(list.get(i).getUsername());
                userModel.setAge(list.get(i).getAge());

                UserModel judgeUser = dao.getUserModelById(list.get(i).getId());

                if(judgeUser!=null){
                    userUpdateList.add(userModel);
                }else{
                    userInsertList.add(userModel);
                }
            }

            if(userUpdateList.size()>0){
                // 批量更新导入数据，一次最多更新1000条
                this.batchUpdateByExcel(userUpdateList, 0, 1000);
            }

            if(userInsertList.size()>0){
                // 批量新增导入数据，一次最多插入1000条
                this.batchInsertByExcel(userInsertList, 0, 1000);
            }

            map.put(ExcelUtils.IS_SUCCESS, true);

        }catch (Exception e){
            map.put(ExcelUtils.IS_SUCCESS, false);
            log.error(e.getMessage(), e);
        }

        return map;
    }

    private void batchUpdateByExcel(List<UserModel> list, int startIndex, int maxSize){
        if (startIndex + maxSize >= list.size()){
            dao.updateByExcel(list.subList(startIndex, list.size()));
        }else{
            dao.updateByExcel(list.subList(startIndex, startIndex + maxSize));
            this.batchUpdateByExcel(list, startIndex + maxSize, maxSize);
        }
    }

    private void batchInsertByExcel(List<UserModel> list, int startIndex, int maxSize){
        if (startIndex + maxSize >= list.size()){
            dao.insertByExcel(list.subList(startIndex, list.size()));
        }else{
            dao.insertByExcel(list.subList(startIndex, startIndex + maxSize));
            this.batchInsertByExcel(list, startIndex + maxSize, maxSize);
        }
    }

}
