package com.muyangren.poidemo.entity;

import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * @author: muyangren
 * @Date: 2023/4/27
 * @Description: 工作情况实体类
 * @Version: 1.0
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class Going {

    /**
     * 工作单位
     */
    private String company;
    /**
     * 地址
     */
    private String address;
    /**
     * 联系电话
     */
    private String phone;

}
