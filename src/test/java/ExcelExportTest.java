import cn.afterturn.easypoi.excel.entity.ExportParams;
import com.lomtom.MyTestApplication;
import com.lomtom.entity.CommodityTemplate;
import com.lomtom.entity.OssProperties;
import com.lomtom.excel.ExcelExportUtils;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author lomtom
 * @date 2021/8/17 9:36
 **/

@SpringBootTest(classes = {MyTestApplication.class})
@RunWith(SpringRunner.class)
public class ExcelExportTest {


    @Autowired
    OssProperties ossProperties;


    List<CommodityTemplate> data = new ArrayList<>();
    {
        data.add(new CommodityTemplate("1","小飞扬","https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic16141250914.JPG",new Date(),"https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic34024334651.JPG"));
        data.add(new CommodityTemplate("2","小道科","https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic3728058748.JPG",new Date(),"https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic22109010312.JPG"));
        data.add(new CommodityTemplate("3","小浩林,","https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic3728058748.JPG",new Date(),"https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic49446558335.JPG"));
    }

    @Test
    public void exportOSS() throws Exception {
        ExportParams params1 = new ExportParams();
        params1.setHeight((short) 28);
        List<CommodityTemplate> list = data;
        String url = ExcelExportUtils.exportExcelByOss("changxing/test/", "ceshi.xls", CommodityTemplate.class, ossProperties, list, params1);
        System.out.println(url);
    }

    @Test
    public void exportLocal() throws Exception {
        ExportParams params1 = new ExportParams();
        params1.setHeight((short) 28);
        List<CommodityTemplate> list = data;
        String url = ExcelExportUtils.exportExcelByLocal("D:\\", "ceshi.xls", CommodityTemplate.class, ossProperties, list, params1);
        System.out.println(url);
    }
}
