package com.poi.demo.entity;

import lombok.Data;
import lombok.ToString;

@Data
@ToString
public class PoiEntity {

    private String id;
    private String breast;
    private String adipocytes;
    private String negative;
    private String staining;
    private String supportive;

}
