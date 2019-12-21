package com.bzc.example.demo;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * 测试控制器
 * @Author baizhicheng
 * @Date 2019/12/14 16:48
 */
@RestController
public class HelloController {

    @RequestMapping("/hello")
    public  String hello(){
        return "Hello SpringBoot!";
    }

}
