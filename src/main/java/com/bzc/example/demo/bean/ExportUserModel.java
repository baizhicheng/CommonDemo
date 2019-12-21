package com.bzc.example.demo.bean;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;

/**
 * @Author baizhicheng
 * @Date 2019/12/21 19:14
 */
@Data
@EqualsAndHashCode(callSuper = false)
@NoArgsConstructor
public class ExportUserModel {

    //ID
    private String id;

    //用户名
    private String username;

    //年龄
    private String age;

}
