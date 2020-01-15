package com.moer.poi.po;

import com.moer.poi.annotation.MyExcelAnno;
import lombok.Data;

import java.io.Serializable;

@Data
public class UserPo implements Serializable {
    private static final long serialVersionUID = 1L;

//    @MyExcelAnno(value="ID")
    private Integer id;
    @MyExcelAnno(value="标题")
    private String title;
    @MyExcelAnno(value="状态",type=1)
    private Integer status;
    @MyExcelAnno(value="用户名")
    private String username;
    @MyExcelAnno(value="性别",type=1,soft=1)
    private String sex;
    @MyExcelAnno(value="所持金额",type=2,format = "{}元")
    private String money;
}
