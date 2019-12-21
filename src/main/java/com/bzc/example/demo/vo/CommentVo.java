package com.bzc.example.demo.vo;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;

/**
 * @Author baizhicheng
 * @Date 2019/12/14 20:16
 */
@Data
@NoArgsConstructor
@EqualsAndHashCode(callSuper = false)
public class CommentVo {

    /**
     * 行下标
     */
    private int row;

    /**
     * 列下标
     */
    private int cell;

    /**
     * 批注信息
     */
    private String msg;

}
