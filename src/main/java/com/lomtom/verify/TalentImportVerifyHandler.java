package com.lomtom.verify;

import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;
import com.lomtom.entity.MemberRecord;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;
import java.util.StringJoiner;

/**
 * @author lomtom
 */
@Component
public class TalentImportVerifyHandler implements IExcelVerifyHandler<MemberRecord> {

    List<String> test1 = new ArrayList<>();


    @Override
    public ExcelVerifyHandlerResult verifyHandler(MemberRecord entity) {
        String asparagusCode = entity.getAsparagusCode();
        StringJoiner msg = new StringJoiner(",");
        if (checkIsDuplicates(asparagusCode)){
            msg.add("MIS编码重复,请核实");
        }
        if (entity.getName() ==null){
            msg.add("名字为空,请核实");
        }
        if (StringUtils.isEmpty(asparagusCode)) {
            msg.add("MIS编码为空,请核实");
        }
        if (msg.length() != 0){
            return new ExcelVerifyHandlerResult(false, msg.toString());
        }
        test1.add(asparagusCode);
        return new ExcelVerifyHandlerResult(true);
    }

    boolean checkIsDuplicates(String asparagusCode){
        return test1.contains(asparagusCode);
    }
}