package com.bzc.example.demo.bean;

import com.bzc.example.demo.annotation.Excel;
import com.bzc.example.demo.annotation.ExcelFied;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;

/**
 * @Author baizhicheng
 * @Date 2019/12/21 23:42
 */
@Data
@EqualsAndHashCode(callSuper = false)
@NoArgsConstructor
@Excel(sheets={"0"}, skip = 0)
public class ImportUserModel {

    //ID
    @ExcelFied(index = 0, isTrim = true)
    private String id;

    //用户名
    @ExcelFied(index = 1, isTrim = true)
    private String username;

    //年龄
    @ExcelFied(index = 2, isTrim = true)
    private String age;

}
