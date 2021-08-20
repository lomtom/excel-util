package com.lomtom.verify;

import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;
import com.lomtom.excel.ExcelImportUtils;

/**
 * @author lomtom
 * @date 2021/8/20 14:00
 **/
public class BaseExcelVerifyHandler<T> implements IExcelVerifyHandler<T> {
    @Override
    public ExcelVerifyHandlerResult verifyHandler(T t) {
        return new ExcelVerifyHandlerResult(!ExcelImportUtils.checkFieldAllNull(t));
    }
}
