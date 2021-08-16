package com.lomtom.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.handler.inter.IExcelDataModel;
import cn.afterturn.easypoi.handler.inter.IExcelModel;
import lombok.Data;



/**
 * @author lomtom
 */
@Data
public class MemberRecord  implements IExcelDataModel, IExcelModel {
    /**
     * 行号
     */
    private Integer rowNum;

    /**
     * 错误消息
     */
    private String errorMsg;
    /**
     * id
     */
    @Excel(name = "账号")
    private String id;

    /**
     * 账号
     */
    @Excel(name = "账号")
    private String account;


    /**
     * 姓名
     */
    @Excel(name = "姓名")
    private String name;

    /**
     * 电话号码
     */
    @Excel(name = "手机号码")
    private String phoneNumber;

    /**
     * 芦笋编码
     */
    @Excel(name = "芦笋编码")
    private String asparagusCode;

    /**
     * 代用编码
     */
    @Excel(name = "代用编码")
    private String substituteCode;

    /**
     * 主体名称
     */
    @Excel(name = "主体名称")
    private String subjectName;

    /**
     * 身份证
     */
    @Excel(name = "身份证号码")
    private String idCard;

    /**
     * 银行卡账号
     */
    @Excel(name = "银行卡号")
    private String bankCard;

    /**
     * 角色ID
     */
    @Excel(name = "角色")
    private String roleId;

    /**
     * 角色名称
     */
    private String roleName;

    /**
     * 角色名称
     */
    private String description;


    /**
     * 删除标志
     */
    private String del;

    /**
     * 创建时间
     */
    private String createTime;

    /**
     * 修改时间
     */
    private String updateTime;

}
