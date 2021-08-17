基于easypoi的工具类封装

表格导入工具类系列：
1. [Easypoi（一）表格导入工具类封装](https://blog.csdn.net/qq_41929184/article/details/119734707)
2. [Easypoi（二）表格工具类使用](https://blog.csdn.net/qq_41929184/article/details/119738265)

注： [工具类及测试github源码](https://github.com/zero028/demo_excel)


### 种类
工具类导入分为六种，分别为：
|函数  | 含义 |
|--|--|
| importExcelByLocalUrlWithImg | 本地文件、带图片  |
|importExcelByLocalUrl | 本地文件、不带图片 |
|importExcelByOssUrlWithImg |OSS文件、带图片  |
|importExcelByOssUrl |OSS文件、不带图片  |
|importExcelByStreamWithImg | 流、带图片  |
|importExcelByStream |流、不带图片  |

注意：若带图片，则需要指定OSS的保存的路径。

```java
    String url1 = "https://demo-excel.oss-cn-hangzhou.aliyuncs.com/测试.xlsx";
    String url2 = "D:\\测试.xlsx";
    String url3 = "D:\\会员档案信息.xls";
```

`OssProperties`为OSS连接参数类，`TalentImportVerifyHandler `为自定义校验类。


```java
@Autowired
OssProperties ossProperties;

@Autowired
TalentImportVerifyHandler talentImportVerifyHandler;

```
## 导入
### 1、调用OSS(带有图片)
```java

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
```
### 2、调用本地(带有图片)
```java
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
```

### 3、调用流（带图片）
```java
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
```

### 4、调用流（不带图片，带校验器）
```java
    /**
     * 调用流
     * @throws Exception Exception
     */
    @Test
    public void importExcelByStream() throws Exception {
        FileInputStream fis = new FileInputStream(url3);
        ImportParams params = new ImportParams();
        ExcelImportResult<MemberRecord> result= ExcelImportUtils.importExcelByStream(fis,MemberRecord.class,params);
        if (result.isVerifyFail()){
            for (MemberRecord entity : result.getFailList()) {
                String msg = "第" + entity.getRowNum() + "行的错误是：" + entity.getErrorMsg();
                System.out.println(msg);
            }
        }
        result.getList().forEach(System.out::println);
    }
```

### 5、调用OSS(不带有图片)
```java
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

```

### 6、调用本地（不带图片）
```java
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
```

### 7、调用流（不带图片，带校验器）
自定义校验器：实现`IExcelVerifyHandler`，编译自己的校验逻辑。
```java
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
```
测试
```java
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
```
结果：
![](https://img-blog.csdnimg.cn/bcf91ca66f4b4aeca5079d0bc734ed13.png)

![](https://img-blog.csdnimg.cn/20487aed4d5a4d539f4b1f194891d3e3.png)


**注意**：
1. 实现` IExcelDataModel`与 `IExcelModel`接口，将不能使用链式编程，即不能在实体类上加`
   @Accessors(chain = true)`。
2. 若项目中还是用了其他excel工具类，请将公有的依赖排除，例如easyexcel
   ```xml
   <dependency>
       <groupId>cn.afterturn</groupId>
       <artifactId>easypoi-base</artifactId>
       <version>2.3.0</version>
       <exclusions>
           <exclusion>
               <groupId>com.google.guava</groupId>
               <artifactId>guava</artifactId>
           </exclusion>
       </exclusions>
   </dependency>
   <dependency>
       <groupId>com.alibaba</groupId>
       <artifactId>easyexcel</artifactId>
       <version>2.2.0-beta2</version>
       <exclusions>
           <exclusion>
               <groupId>org.apache.poi</groupId>
               <artifactId>poi-ooxml</artifactId>
           </exclusion>
           <exclusion>
               <groupId>org.apache.poi</groupId>
               <artifactId>poi-ooxml-schemas</artifactId>
           </exclusion>
       </exclusions>
   </dependency>
   ```
3. 若使用了`easyexcel`，在引入easypoi并且排除共有依赖导致easyexcel的部分不可用，请升级easypoi的版本。

```xml
 <dependency>
     <groupId>cn.afterturn</groupId>
     <artifactId>easypoi-base</artifactId>
     <version>4.1.3</version>
 </dependency>
```
## 导出
### 1、导出到本地
需要指定文件保存的路径与文件名
```java
List<CommodityTemplate> data = new ArrayList<>();
{
    data.add(new CommodityTemplate("1","小飞扬","https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic16141250914.JPG",new Date(),"https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic34024334651.JPG"));
    data.add(new CommodityTemplate("2","小道科","https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic3728058748.JPG",new Date(),"https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic22109010312.JPG"));
    data.add(new CommodityTemplate("3","小浩林,","https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic3728058748.JPG",new Date(),"https://demo-excel.oss-cn-hangzhou.aliyuncs.com/changxing/test/img/imggggpic49446558335.JPG"));
}

@Test
public void exportLocal() throws Exception {
    ExportParams params1 = new ExportParams();
    params1.setHeight((short) 28);
    List<CommodityTemplate> list = data;
    String url = ExcelExportUtils.exportExcelByLocal("D:\\", "ceshi.xls", CommodityTemplate.class, ossProperties, list, params1);
    System.out.println(url);
}
```

### 2、导出到OSS
需要指定文件保存的路径（不包括域名）与文件名
```java
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
```
![](https://img-blog.csdnimg.cn/94c5095d8b2041d9b9b65f3a34a5f064.png#pic_center)
