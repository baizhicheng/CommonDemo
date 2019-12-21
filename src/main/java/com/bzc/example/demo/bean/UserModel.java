package com.bzc.example.demo.bean;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;

import javax.annotation.sql.DataSourceDefinition;

/**
 * @Author baizhicheng
 * @Date 2019/12/14 18:48
 */
@Data
@EqualsAndHashCode(callSuper = false)
//生成一个无参数的构造方法
@NoArgsConstructor
public class UserModel {

    //ID
    private String id;

    //用户名
    private String username;

    //年龄
    private String age;

}

