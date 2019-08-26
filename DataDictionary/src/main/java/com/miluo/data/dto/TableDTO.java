package com.miluo.data.dto;

import lombok.Data;

import java.util.List;

@Data
public class TableDTO {

    private String tableName;
    private String tableRemark;
    private List<FieldDTO> fields;

}
