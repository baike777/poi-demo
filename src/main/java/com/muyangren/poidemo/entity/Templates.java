package com.muyangren.poidemo.entity;

import lombok.Data;
import lombok.EqualsAndHashCode;

import java.io.Serializable;

/**
 * @author: muyangren
 * @Date: 2023/1/16
 * @Description: 基础数据
 * @Version: 1.0
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class Templates implements Serializable {
    private static final long serialVersionUID = 1L;

    /**
     * 姓名
     */
    private String name;
    /**
     * 别名
     */
    private String aliases;
    /**
     * 部门
     */
    private String deptName;
    /**
     * 性别
     */
    private String sexName;
    /**
     * 民族
     */
    private String peoples;
    /**
     * 出生日期
     */
    private String birth;
    /**
     * 文化程度
     */
    private String culture;
    // 。。。。其他字段省略 太累了


}
