package com.moer.poi.utils;

import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
import com.moer.poi.annotation.MyExcelAnno;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;

import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

@Data
public class PoiExcelUtils {

    final static int startIndex = 0; // 标题行下标,默认设成从0开始
    private String title = "excel";

    /**
     * 初始化workbook
     * @return
     */
    public Workbook initWorkbook() {
        Workbook workbook = new HSSFWorkbook();
        workbook.createSheet(title);
        // 创建sheet，写入枚举项
        workbook.createSheet("hiddenSheet");
        return workbook;
    }

    /**
     * 导出excel数据 可变换枚举
     *
     * @param pojoClass
     * @return
     */
    public Workbook exportExcelDatg(Class<?> pojoClass, List<Object> listObj, Object[] args) {
        Workbook workbook = initWorkbook();
        initExcelTitle(workbook, pojoClass);
        initExcelData(workbook, listObj,args);
        return workbook;
    }

    /**
     * 导出excel数据
     *
     * @param pojoClass
     * @return
     */
    public Workbook exportExcelDatg(Class<?> pojoClass, List<Object> listObj) {
        Workbook workbook = initWorkbook();
        initExcelTitle(workbook, pojoClass);
        initExcelData(workbook, listObj,null);
        return workbook;
    }


    /**
     * 导出excel模板 带下拉框
     *
     * @param pojoClass
     * @param arg
     * @return
     */
    public Workbook exportExcel(Class<?> pojoClass, Object[] arg) {
        Workbook workbook = initWorkbook();
        initExcelTitle(workbook, pojoClass);
        initSelectCol(workbook, pojoClass, arg);
        return workbook;
    }

    /**
     * 导出excel模板
     *
     * @param pojoClass
     * @return
     */
    public Workbook exportExcel(Class<?> pojoClass) {
        Workbook workbook = initWorkbook();
        initExcelTitle(workbook, pojoClass);
        return workbook;
    }

    /**
     * 初始化导出excel的标题行
     * @param workbook
     * @param pojoClass po的class
     */
    public void initExcelTitle(Workbook workbook, Class<?> pojoClass) {
        int excelColLength = 0; // 列数
        int cellIndex = 0; // 已经创建的列数下标
        Field[] fields = pojoClass.getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            // 设置是否允许访问，不是修改原来的访问权限修饰词。
            fields[i].setAccessible(true);
            MyExcelAnno excelAnno = fields[i].getAnnotation(MyExcelAnno.class);
            if (excelAnno != null) {
                excelColLength++;
            }
        }

        Sheet sheet = workbook.getSheet(title);
        // 设置文本框默认格式为文本
        for (int i = 0; i < excelColLength; i++) {
            CellStyle cellStyle = workbook.createCellStyle();
            DataFormat dataFormat = workbook.createDataFormat();
            cellStyle.setDataFormat(dataFormat.getFormat("@"));
            sheet.setDefaultColumnStyle(i, cellStyle);
        }

        // 设置默认列宽度
//		sheet.setDefaultColumnWidth(15);

        Row row = sheet.createRow(startIndex);

        for (int i = 0; i < fields.length; i++) {
            // 设置是否允许访问，不是修改原来的访问权限修饰词。
            fields[i].setAccessible(true);
            MyExcelAnno excelAnno = fields[i].getAnnotation(MyExcelAnno.class);
            if (excelAnno != null) {
                row.createCell(cellIndex).setCellValue(excelAnno.value());
                sheet.setColumnWidth(cellIndex, (excelAnno.value().getBytes().length + 8) * 256);
                cellIndex++;
            }
        }
    }


    /**
     * 写入excel数据
     *
     * @param workbook
     * @param listObj 需要导出的list集合
     * @param args 为需要进行枚举转换的map<Integer,String>格式,根据po中的soft顺序,放入object[]集合中
     */
    public void initExcelData(Workbook workbook, List<Object> listObj, Object[] args) {
        int cellIndex = 0; // 已经创建的列数下标
        int rowIndex = startIndex + 1; // 行数下标
        Sheet sheet = workbook.getSheet(title);
        for (Object o : listObj) {
            Field[] fields = o.getClass().getDeclaredFields();
            Row row = sheet.createRow(rowIndex);
            for (Field field : fields) {
                MyExcelAnno excelAnno = field.getAnnotation(MyExcelAnno.class);
                if (excelAnno != null) {
                    //获取实体类中的值
                    String value = (String) ReflectUtil.getFieldValue(o, field);;
                    if (excelAnno.type() == 1) {//进行枚举转换
                        Map<Integer, String> arg = (Map<Integer, String>) args[excelAnno.soft()];
                        for (Map.Entry<Integer, String> argEntry : arg.entrySet()) {
                            if (argEntry.getKey().toString().equals(value)) {
                                row.createCell(cellIndex).setCellValue(argEntry.getValue());
                            }
                        }
                    }else if(excelAnno.type() == 2) {//进行格式化
                        row.createCell(cellIndex).setCellValue(StrUtil.format(excelAnno.format(),value));
                    }else {
                        row.createCell(cellIndex).setCellValue(value);
                    }
                    cellIndex++;
                }
            }
            rowIndex++;
            cellIndex = 0;
        }
    }

    /**
     * 导出模板-设置下拉框值
     * @param workbook
     * @param pojoClass
     * @param arg
     */
    public void initSelectCol(Workbook workbook, Class<?> pojoClass, Object[] arg) {
        int tmpArgIndex = 0;
        int cellIndex = 0;
        Field[] fields = pojoClass.getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            //设置是否允许访问，不是修改原来的访问权限修饰词。
            fields[i].setAccessible(true);
            MyExcelAnno excelAnno = fields[i].getAnnotation(MyExcelAnno.class);
            if (excelAnno != null) {
                //如果是1进行下拉框赋值
                if (excelAnno.type() == 1) {
                    XSSFSetDropDownAndHidden(workbook, cellIndex, tmpArgIndex, (String[]) arg[tmpArgIndex]);
                    tmpArgIndex++;
                }
                cellIndex++;
            }
        }
    }

    /**
     * 使用createFormulaListConstraint实现下拉框
     *
     * @param workbook
     * @param formulaString 下拉框数据
     * @param colNum        第几列插入
     * @param index         是第几个下拉框格式数据,从0开始计算
     * @return
     */
    public void XSSFSetDropDownAndHidden(Workbook workbook, Integer colNum, Integer index, String[] formulaString) {
        String az = "ABCDEFGHIJKLMNPQRSTUVWXYZ";
        char letter = az.charAt(index);

        Sheet sheet = workbook.getSheet("录入模板");
        Sheet hideSheet = workbook.getSheet("hiddenSheet");

        for (int i = 0; i < formulaString.length; i++) {
            Row row = hideSheet.getRow(i);
            if (row != null) {
                row.createCell(1).setCellValue(formulaString[i]);
            } else {
                hideSheet.createRow(i).createCell(0).setCellValue(formulaString[i]);
            }
        }

        // 创建名称，可被其他单元格引用
        Name category1Name = workbook.createName();
        category1Name.setNameName("hidden" + index);
        // 设置名称引用的公式
        // 使用像'A1：B1'这样的相对值会导致在Microsoft Excel中使用工作簿时名称所指向的单元格的意外移动，
        // 通常使用绝对引用，例如'$A$1:$B$1'可以避免这种情况。
        // 参考： http://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Name.html
        category1Name.setRefersToFormula("hiddenSheet!" + "$" + letter + "$1:$" + letter + "$" + formulaString.length);
        // 获取上文名称内数据
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createFormulaListConstraint("hidden" + index);
        // 设置下拉框位置
        CellRangeAddressList addressList = new CellRangeAddressList(startIndex + 1, 65535, colNum, colNum);
        DataValidation dataValidation = helper.createValidation(constraint, addressList);
        // 处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            // 数据校验
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        // 作用在目标sheet上
        sheet.addValidationData(dataValidation);
        // 设置hiddenSheet隐藏
        workbook.setSheetHidden(1, true);
    }



    /**
     * 读取excelfile返回maplist
     * @return
     * @throws IOException
     */
    public List<Map<String, Object>> getMaps(InputStream stream, Class<?> pojoClass) throws IOException {
        //输入输出流
//		InputStream is = new FileInputStream("d:\\testDropDown.xlsx");
        //创建工作空间
        Workbook wb = WorkbookFactory.create(stream);
        //获取工作表
        Sheet sheet = wb.getSheetAt(0);//获取第一个工作表
        //获取标题,根据反射转换成实体类变量名
        Row row = sheet.getRow(startIndex);
        List<Map<String, Object>> readMapList = new ArrayList<>();
        String[] titleArg = new String[row.getLastCellNum()];
        for (int i = 0; i < row.getLastCellNum(); i++) {
            titleArg[i] = row.getCell(i).toString();
        }
        //英文key
        String[] keyArgs = changeKeyByArray(pojoClass, titleArg);
        //获取数据，放入map中
        for (int i = startIndex + 1; i < sheet.getLastRowNum()+1; i++) {
            Row row1 = sheet.getRow(i);
            Map<String, Object> dtoMap = new HashMap<>();
            for (int j = 0; j < row1.getLastCellNum(); j++) {
                Cell cell = row1.getCell(j);
                if(cell != null){
                    cell.setCellType(CellType.STRING);
                    dtoMap.put(keyArgs[j], cell.getStringCellValue());
                }
            }
            readMapList.add(dtoMap);
        }
        return readMapList;
    }

    /**
     * 把从excel读取的中文key转换为英文属性名
     *
     * @param pojoClass
     * @param args
     * @return
     */
    public static String[] changeKeyByArray(Class<?> pojoClass, String[] args) {
        String[] reKeys = new String[args.length];
        Field[] fields = pojoClass.getDeclaredFields();
        Arrays.asList(fields).forEach(field -> {
            MyExcelAnno excelAnno = field.getAnnotation(MyExcelAnno.class);
            if (excelAnno != null) {
                //如果中文名一样的话,把数据存入新map中
                for (int i = 0; i < args.length; i++) {
                    if (args[i].equals(excelAnno.value())) {
                        reKeys[i] = field.getName();
                    }
                }
            }
        });
        return reKeys;
    }

    /**
     * 根据反射吧传进来的MAP中下拉框从中文转换成数字枚举
     *
     * @param excelMap  从excel读取的map数据
     * @param args 都是Map<Integer, String>类型 例如1,男|2,女
     * @return
     */
    public static Map<String, Object> changeKeyByMap(Class<?> pojoClass, Map<String, Object> excelMap, Object[] args) {
        Map<String, Object> rtMap = new HashMap<>();
        Field[] fields = pojoClass.getDeclaredFields();
        //遍历读取的map数据
        for (Map.Entry<String, Object> entry : excelMap.entrySet()) {
            String k = entry.getKey();
            Object v = entry.getValue();

            //遍历字段名进行类型判断
            for (Field field : fields) {
                MyExcelAnno excelAnno = field.getAnnotation(MyExcelAnno.class);
                if (excelAnno != null) {
                    //遍历map进行下拉框类型判断,如果字段名跟map的key相等，进行判断
                    if(field.getName().equals(k)){
                        if (excelAnno.type() == 1) {
                            Map<Integer, String> arg = (Map<Integer, String>) args[excelAnno.soft()];
                            for (Map.Entry<Integer, String> argEntry : arg.entrySet()) {
                                Integer q = argEntry.getKey();
                                Object w = argEntry.getValue();
                                if (w.toString().equals(v)) {//如果内容相等进行转换中文为int
                                    rtMap.put(k, q);
                                }
                            }
                        } else {
                            //如果相等那么进行类型判断
                            if (field.getName().equals(k)) {
                                rtMap.put(k, v);
                            }
                        }
                    }
                }
            }
        }
        return rtMap;
    }
}
