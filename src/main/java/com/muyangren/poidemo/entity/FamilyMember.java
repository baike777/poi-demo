package com.muyangren.poidemo.entity;

import lombok.Data;
import lombok.EqualsAndHashCode;

import java.io.Serializable;

/**
 * @author: muyangren
 * @Date: 2023/4/27
 * @Description: 家庭成员实体类
 * @Version: 1.0
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class FamilyMember implements Serializable {
    /**
     * 关系
     */
    private String relationship;
    /**
     * 姓名
     */
    private String name;
    /**
     * 职位
     */
    private String position;

}
