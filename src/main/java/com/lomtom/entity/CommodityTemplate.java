package com.lomtom.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.handler.inter.IExcelDataModel;
import cn.afterturn.easypoi.handler.inter.IExcelModel;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;

import java.util.Date;

/**
 * @author Feyoung
 * @description：商品入库模板
 * @date 2021/8/4 14:42
 */
@Data
public class CommodityTemplate implements IExcelDataModel, IExcelModel {
    /**
     * 行号
     */
    private Integer rowNum;

    /**
     * 错误消息
     */
    private String errorMsg;

    @Excel(name = "编号")
    private String id;

    /**
     * 商品名字
     */
    @Excel(name = "名字")
    private String name;

    /**
     * 商品图片
     */
    @Excel(name = "图片", type = 2, savePath = "\\imgggg",width = 20)
    private String imageUrl;

    /**
     * 创建时间
     */
    @Excel(name = "日期",exportFormat = "yyyy-MM-dd")
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date createTime;

    @Excel(name = "头像", type = 2, savePath = "\\imgggg",width = 20)
    private String img;

    public CommodityTemplate(String id, String name, String imageUrl, Date createTime, String img) {
        this.id = id;
        this.name = name;
        this.imageUrl = imageUrl;
        this.createTime = createTime;
        this.img = img;
    }
}
