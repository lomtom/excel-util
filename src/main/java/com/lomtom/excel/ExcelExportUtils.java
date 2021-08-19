package com.lomtom.excel;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import com.aliyun.oss.OSS;
import com.aliyun.oss.OSSClientBuilder;
import com.lomtom.entity.OssProperties;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.util.List;

/**
 * @author lomtom
 * @date 2021/8/17 10:33
 **/
public class ExcelExportUtils {

    public static <T> String exportExcelToOss(String rootPath, String fileName, Class<T> tClass, OssProperties ossProperties, List<T> data, ExportParams params) throws Exception {
        return exportExcelToOss(rootPath, fileName,tClass, ossProperties.getBucket(), ossProperties.getRegionId(), ossProperties.getAk(), ossProperties.getSk(),data, params);
    }


    public static <T> String exportExcelToLocal(String rootPath, String fileName, Class<T> tClass, List<T> data, ExportParams params) throws Exception {
        String fullPath = rootPath + fileName;
        Workbook sheets = ExcelExportUtil.exportExcel(params, tClass, data);
        FileOutputStream stream = new FileOutputStream(fullPath);
        sheets.write(stream);
        stream.close();
        return fullPath;
    }


    public static <T> byte[] exportExcelToOutPutStream(Class<T> tClass, List<T> data, ExportParams params) throws Exception {
        Workbook sheets = ExcelExportUtil.exportExcel(params, tClass, data);
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        sheets.write(stream);
        return stream.toByteArray();
    }


    private static <T> String exportExcelToOss(String rootPath, String fileName, Class<T> tClass, String bucket, String regionId, String accessKeyId, String accessKeySecret, List<T> data, ExportParams params) throws Exception {
        String head = "https://oss-" + regionId + ".aliyuncs.com";
        String root = "https://" + bucket + ".oss-" + regionId + ".aliyuncs.com/";
        String fullPath = rootPath + fileName;
        Workbook sheets = ExcelExportUtil.exportExcel(params, tClass, data);
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        sheets.write(bos);
        byte[] bytes = bos.toByteArray();
        OSS ossClient = new OSSClientBuilder().build(head, accessKeyId, accessKeySecret);
        ossClient.putObject(bucket, fullPath, new ByteArrayInputStream(bytes));
        bos.close();
        ossClient.shutdown();
        return root + fullPath;
    }
}
