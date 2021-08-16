import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import com.lomtom.MyTestApplication;
import com.lomtom.entity.CommodityTemplate;
import com.lomtom.entity.MemberRecord;
import com.lomtom.entity.OssProperties;
import com.lomtom.excel.ExcelImportUtils;
import com.lomtom.verify.TalentImportVerifyHandler;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.FileInputStream;

/**
 * @author lomtom
 * @date 2021/8/16 11:13
 **/
@SpringBootTest(classes = {MyTestApplication.class})
@RunWith(SpringRunner.class)
public class ExcelTest {

    String url1 = "https://demo-excel.oss-cn-hangzhou.aliyuncs.com/测试.xlsx";
    String url2 = "D:\\测试.xlsx";
    String url3 = "D:\\会员档案信息.xls";

    @Autowired
    OssProperties ossProperties;

    @Autowired
    TalentImportVerifyHandler talentImportVerifyHandler;



    /**
     * 调用OSS(带有图片)
     * @throws Exception Exception
     */
    @Test
    public void importExcelByOssUrlWithImg() throws Exception {
        ImportParams params = new ImportParams();
        ExcelImportResult<CommodityTemplate> result= ExcelImportUtils.importExcelByOssUrlWithImg(url1,
                CommodityTemplate.class,ossProperties,"changxing/test/",params);
        if (result.isVerifyFail()){
            for (CommodityTemplate entity : result.getFailList()) {
                String msg = "第" + entity.getRowNum() + "行的错误是：" + entity.getErrorMsg();
                System.out.println(msg);
            }
        }
        result.getList().forEach(System.out::println);
    }

    /**
     * 调用本地(带有图片)
     * @throws Exception Exception
     */
    @Test
    public void importExcelByLocalUrlWithImg() throws Exception {
        ImportParams params = new ImportParams();
        ExcelImportResult<CommodityTemplate> result= ExcelImportUtils.importExcelByLocalUrlWithImg(url2,
                CommodityTemplate.class,ossProperties,"changxing/test/",params);
        if (result.isVerifyFail()){
            for (CommodityTemplate entity : result.getFailList()) {
                String msg = "第" + entity.getRowNum() + "行的错误是：" + entity.getErrorMsg();
                System.out.println(msg);
            }
        }
        result.getList().forEach(System.out::println);
    }

    /**
     * 调用流（带图片）
     * @throws Exception Exception
     */
    @Test
    public void importExcelByStreamWithImg() throws Exception {
        FileInputStream fis = new FileInputStream(url2);
        ImportParams params = new ImportParams();
        ExcelImportResult<CommodityTemplate> result= ExcelImportUtils.importExcelByStreamWithImg(fis,CommodityTemplate.class,ossProperties,"changxing/test/",params);
        if (result.isVerifyFail()){
            for (CommodityTemplate entity : result.getFailList()) {
                String msg = "第" + entity.getRowNum() + "行的错误是：" + entity.getErrorMsg();
                System.out.println(msg);
            }
        }
        result.getList().forEach(System.out::println);
    }

    /**
     * 调用流（不带图片，带校验器）
     * @throws Exception Exception
     */
    @Test
    public void importExcelByStream() throws Exception {
        FileInputStream fis = new FileInputStream(url3);
        ImportParams params = new ImportParams();
        //自定义校验器
        params.setVerifyHandler(talentImportVerifyHandler);
        ExcelImportResult<MemberRecord> result= ExcelImportUtils.importExcelByStream(fis,MemberRecord.class,params);
        if (result.isVerifyFail()){
            for (MemberRecord entity : result.getFailList()) {
                String msg = "第" + entity.getRowNum() + "行的错误是：" + entity.getErrorMsg();
                System.out.println(msg);
            }
        }
        result.getList().forEach(System.out::println);
    }

    /**
     * 调用OSS(不带有图片)
     * @throws Exception Exception
     */
    @Test
    public void importExcelByOssUrl() throws Exception {
        ImportParams params = new ImportParams();
        ExcelImportResult<CommodityTemplate> result= ExcelImportUtils.importExcelByOssUrl(url1,
                CommodityTemplate.class,ossProperties,params);
        if (result.isVerifyFail()){
            for (CommodityTemplate entity : result.getFailList()) {
                String msg = "第" + entity.getRowNum() + "行的错误是：" + entity.getErrorMsg();
                System.out.println(msg);
            }
        }
        result.getList().forEach(System.out::println);
    }


    /**
     * 调用本地（不带图片）
     * @throws Exception Exception
     */
    @Test
    public void importExcelByLocalUrl() throws Exception {
        ImportParams params = new ImportParams();
        ExcelImportResult<MemberRecord> result= ExcelImportUtils.importExcelByLocalUrl(url3,MemberRecord.class,params);
        if (result.isVerifyFail()){
            for (MemberRecord entity : result.getFailList()) {
                String msg = "第" + entity.getRowNum() + "行的错误是：" + entity.getErrorMsg();
                System.out.println(msg);
            }
        }
        result.getList().forEach(System.out::println);
    }

}
