package com.miluo.data.dto;

import com.github.crab2died.annotation.ExcelField;
import lombok.Data;

@Data
public class FieldDTO {

    @ExcelField(title = "名称", order = 1)
    private String columnName;

    @ExcelField(title = "代码", order = 2)
    private String columnCode;

    @ExcelField(title = "数据类型", order = 3)
    private String dataType;

    @ExcelField(title = "长度", order = 4)
    private int size;

    @ExcelField(title = "注释", order = 5)
    private String comment;

    @ExcelField(title = "PK", order = 6)
    private String PK;

    @ExcelField(title = "FK", order = 7)
    private String FK;

    @ExcelField(title = "UK", order = 8)
    private String UK;
}
