package com.lomtom.entity;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Data
@Component
@ConfigurationProperties(prefix = "oss")
public class OssProperties {
    private String bucket;
    private String regionId;
    private String ak;
    private String sk;
    private long maxLength = 1048576000;
    private long outTime = 30*60*1000;
}
