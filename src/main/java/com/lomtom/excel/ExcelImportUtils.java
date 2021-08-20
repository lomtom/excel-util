package com.lomtom.excel;

import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import com.aliyun.oss.OSS;
import com.aliyun.oss.OSSClientBuilder;
import com.aliyun.oss.model.OSSObject;
import com.lomtom.entity.OssProperties;
import org.springframework.util.ObjectUtils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.List;

/**
 * @author lomtom
 * @date 2021/8/14 15:44
 **/

public class ExcelImportUtils {

    public static <T> ExcelImportResult<T> importExcelByLocalUrlWithImg(String url, Class<T> tClass, OssProperties ossProperties, String rootPath, ImportParams importParams) throws Exception {
        return importExcelByLocalUrlWithImg(url, tClass, ossProperties.getBucket(), ossProperties.getRegionId(), ossProperties.getAk(), ossProperties.getSk(), rootPath, importParams);
    }

    public static <T> ExcelImportResult<T> importExcelByLocalUrl(String url, Class<T> tClass, ImportParams importParams) throws Exception {
        InputStream in = new FileInputStream(url);
        return importExcelByStream(in, tClass, importParams);
    }

    public static <T> ExcelImportResult<T> importExcelByOssUrlWithImg(String url, Class<T> tClass, OssProperties ossProperties, String rootPath, ImportParams importParams) throws Exception {
        return importExcelByOssUrlWithImg(url, tClass, ossProperties.getBucket(), ossProperties.getRegionId(), ossProperties.getAk(), ossProperties.getSk(), rootPath, importParams);
    }


    public static <T> ExcelImportResult<T> importExcelByOssUrl(String url, Class<T> tClass, OssProperties ossProperties, ImportParams importParams) throws Exception {
        return importExcelByOssUrl(url, tClass, ossProperties.getBucket(), ossProperties.getRegionId(), ossProperties.getAk(), ossProperties.getSk(), importParams);
    }

    public static <T> ExcelImportResult<T> importExcelByStreamWithImg(InputStream in, Class<T> tClass, OssProperties ossProperties, String rootPath, ImportParams importParams) throws Exception {
        return importExcelByStreamWithImg(in, tClass, ossProperties.getBucket(), ossProperties.getRegionId(), ossProperties.getAk(), ossProperties.getSk(), rootPath, importParams);
    }

    public static <T> ExcelImportResult<T> importExcelByStream(InputStream in, Class<T> tClass, ImportParams importParams) throws Exception {
        return getResults(in, tClass, importParams);
    }

    /**
     * 通过流导入
     *
     * @param in              流
     * @param tClass          需要装入的类型
     * @param bucket          bucket
     * @param regionId        regionId
     * @param accessKeyId     accessKeyId
     * @param accessKeySecret accessKeySecret
     * @param rootPath        图片保存的根路径
     * @param <T>             需要装入的类型
     * @return 结果
     * @throws Exception Exception
     */
    private static <T> ExcelImportResult<T> importExcelByStreamWithImg(InputStream in, Class<T> tClass, String bucket, String regionId, String accessKeyId, String accessKeySecret, String rootPath, ImportParams importParams) throws Exception {
        String endpoint = "https://oss-" + regionId + ".aliyuncs.com";
        OSS ossClient = new OSSClientBuilder().build(endpoint, accessKeyId, accessKeySecret);
        ExcelImportResult<T> results = getResults(in, tClass, importParams);
        uploadImg(results.getList(), ossClient, bucket, rootPath, regionId);
        ossClient.shutdown();
        return results;
    }

    /**
     * 通过本地文件导入
     *
     * @param url             本地文件地址
     * @param tClass          需要装入的类型
     * @param bucket          bucket
     * @param regionId        regionId
     * @param accessKeyId     accessKeyId
     * @param accessKeySecret accessKeySecret
     * @param rootPath        图片保存的根路径
     * @param <T>             需要装入的类型
     * @return 结果
     * @throws Exception Exception
     */
    private static <T> ExcelImportResult<T> importExcelByLocalUrlWithImg(String url, Class<T> tClass, String bucket, String regionId, String accessKeyId, String accessKeySecret, String rootPath, ImportParams importParams) throws Exception {
        InputStream in = new FileInputStream(url);
        return importExcelByStreamWithImg(in, tClass, bucket, regionId, accessKeyId, accessKeySecret, rootPath, importParams);
    }

    /**
     * 通过OSS导入
     *
     * @param url             OSS文件地址（全路径）
     * @param tClass          需要装入的类型
     * @param bucket          bucket
     * @param regionId        regionId
     * @param accessKeyId     accessKeyId
     * @param accessKeySecret accessKeySecret
     * @param rootPath        图片保存的根路径
     * @param <T>             需要装入的类型
     * @return 结果
     * @throws Exception Exception
     */
    private static <T> ExcelImportResult<T> importExcelByOssUrlWithImg(String url, Class<T> tClass, String bucket, String regionId, String accessKeyId, String accessKeySecret, String rootPath, ImportParams importParams) throws Exception {
        String head = "https://oss-" + regionId + ".aliyuncs.com";
        String root = "https://" + bucket + ".oss-" + regionId + ".aliyuncs.com/";
        url = url.replace(root, "");
        OSS ossClient = new OSSClientBuilder().build(head,
                accessKeyId, accessKeySecret);
        OSSObject ossObject = ossClient.getObject(bucket, url);
        InputStream in = ossObject.getObjectContent();
        ExcelImportResult<T> results = getResults(in, tClass, importParams);
        uploadImg(results.getList(), ossClient, bucket, rootPath, regionId);
        ossClient.shutdown();
        return results;
    }

    private static <T> ExcelImportResult<T> importExcelByOssUrl(String url, Class<T> tClass, String bucket, String regionId, String accessKeyId, String accessKeySecret, ImportParams importParams) throws Exception {
        String head = "https://oss-" + regionId + ".aliyuncs.com";
        String root = "https://" + bucket + ".oss-" + regionId + ".aliyuncs.com/";
        url = url.replace(root, "");
        OSS ossClient = new OSSClientBuilder().build(head,
                accessKeyId, accessKeySecret);
        OSSObject ossObject = ossClient.getObject(bucket, url);
        InputStream in = ossObject.getObjectContent();
        ExcelImportResult<T> results = getResults(in, tClass, importParams);
        ossClient.shutdown();
        return results;
    }

    private static <T> ExcelImportResult<T> getResults(InputStream in, Class<T> tClass, ImportParams importParams) throws Exception {
        return ExcelImportUtil.importExcelMore(in, tClass, importParams);
    }


    private static <T> void uploadImg(List<T> list, OSS ossClient, String bucket, String rootPath, String regionId) {
        list.forEach(obj -> {
            List<Field> fieldList = Arrays.asList(obj.getClass().getDeclaredFields());
            fieldList.forEach(item -> {
                boolean annotationPresent = item.isAnnotationPresent(Excel.class);
                if (annotationPresent) {
                    Excel anno = item.getAnnotation(Excel.class);
                    if (anno.type() == 2) {
                        try {
                            String path = getFieldValueByName(item.getName(), obj);
                            InputStream inputStream = new FileInputStream(path);
                            String ossFileName = path.replace("\\", "");
                            ossClient.putObject(bucket, rootPath + ossFileName, inputStream);
                            String ossPath = "https://" + bucket + "." + "oss-" + regionId + ".aliyuncs.com" + "/" + rootPath + ossFileName;
                            item.setAccessible(true);
                            item.set(obj, ossPath);
                        } catch (FileNotFoundException | InvocationTargetException | NoSuchMethodException | IllegalAccessException e) {
                            e.printStackTrace();
                        }
                    }
                }
            });
        });
    }

    private static String getFieldValueByName(String fieldName, Object o) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        String firstLetter = fieldName.substring(0, 1).toUpperCase();
        String getter = "get" + firstLetter + fieldName.substring(1);
        Method method = o.getClass().getMethod(getter);
        Object value = method.invoke(o);
        return value.toString();
    }


    /**
     * check nullable
     * @param o o
     * @return result
     */
    public static boolean checkFieldAllNull(Object o){
        try{
            for(Field field:o.getClass().getDeclaredFields()){
                if ("rowNum".equals(field.getName()) || "errorMsg".equals(field.getName())){
                    continue;
                }
                field.setAccessible(true);
                Object object = field.get(o);
                if(!ObjectUtils.isEmpty(object)){
                    return false;
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return true;
    }
}
