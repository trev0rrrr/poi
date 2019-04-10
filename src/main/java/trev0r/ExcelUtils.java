/*
package trev0r;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.binary.XSSFBUtils;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

*/
/**
 * 日期格式序号14  XssfCellStyle1.setDataFormat(14);
 * 设置cell编码解决中文高位字节截断//chCell.setEncoding(HSSFCell.ENCODING_UTF_16);
 *//*

public class ExcelUtils {
    static final String tbd = "未知excel输入类型、请联系管理员";
    static final String tbe = "有未实现功能、";
    //static final int max=Integer.MAX_VALUE;
    static final int maxRowToValidate = 1000;
    static final DateFormat dateFormat = new SimpleDateFormat("yyyy/M/d");//m/d/yy  yyyy/MM/dd


    public static Object dataProcess(Cell cell, Object type) throws Exception {
        */
/*org.apache.poi.ss.usermodel.BuiltinFormats
        short format = cell.getCellStyle().getDataFormat();*//*

        if (cell == null) return null;
        if (type == null) {
            */
/*int i = cell.getCellType();
            if(Cell.CELL_TYPE_NUMERIC == cell.getCellType())
                return Double.toString(cell.getNumericCellValue());
            else if(Cell.CELL_TYPE_STRING == cell.getCellType())
                return cell.getStringCellValue();
            else if(Cell.CELL_TYPE_BLANK == cell.getCellType())
                return null;
            else*//*

            throw new Exception(tbd + "type 为空 异常");
        } else if (type instanceof Integer) {
            Integer i = (Integer) type;
            switch (i) {
                case 1:
                    try {
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_NUMERIC:
                                return Integer.valueOf((int) cell.getNumericCellValue());
                            case Cell.CELL_TYPE_STRING:
                                return Integer.valueOf(cell.getStringCellValue());
                            default:
                                throw new Exception("");
                        }

                    } catch (Exception e) {
                        throw new Exception("未知异常-1");
                    }
                case 14:
                    try {
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_NUMERIC:
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    return cell.getDateCellValue();
                                }
                                return null;//不合理的输入 直接将日期置空
                            case Cell.CELL_TYPE_STRING:
                                return dateFormat.parse(cell.getStringCellValue());
                            case Cell.CELL_TYPE_BLANK:
                                return null;
                            default:
                                throw new Exception("未知的日期输入类型");
                        }
                    } catch (Exception e) {
                        throw new Exception(e.getCause().toString());
                    }

                case 49:
                    try {
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_NUMERIC: {
                                throw new Exception("输入数字怎么处理");
                                //return null;//不合理的输入 直接将日期置空
                            }
                            case Cell.CELL_TYPE_STRING:
                                return cell.getStringCellValue();
                            case Cell.CELL_TYPE_BLANK:
                                return null;
                            default:
                                throw new Exception("未知的String输入类型");
                        }
                    } catch (Exception e) {
                        throw new Exception(e.getCause().toString());
                    }
                    */
/*return (cell.getCellType()==Cell.CELL_TYPE_STRING)
                            ? cell.getStringCellValue()
                            : ((Integer)(((Double)cell.getNumericCellValue()).intValue())).toString();*//*

                default:
                    System.err.println("需要处理");
                    throw new Exception(tbe + "待支持其他类型");
            }
        } else
            throw new Exception(tbe + "自定义格式化输入类型的支持、还没有实现");
    }

    public static void main(String[] args) throws Exception {
        String template = "chip-mdm-parent\\chip-mdm-admin\\src\\main\\resources\\excelTemplate\\人员信息模板" + Math.random() + ".xlsx";
        String importPath = "chip-mdm-parent\\chip-mdm-admin\\src\\main\\resources\\excelTemplate\\人员信息模板 - 副本.xlsx";

        String template2 = "chip-mdm-parent\\chip-mdm-admin\\src\\main\\resources\\excelTemplate\\人员信息模板" + Math.random() + ".xlsx";
        String importPath2 = "C:\\Users\\trev0r\\desktop\\test\\人员信息模板1.xlsx";
        //todo
        List<List<ExcelConfig>> list = null;
        String arr[] = null;

        generateTemplate(template, list, null);
        //List result = importFromExcel(new FileInputStream(new File(importPath2)),list,arr);
        ArrayList<List<Map<String, Object>>> sources = new ArrayList();
        ArrayList<Map<String, Object>> source = new ArrayList();
        Map<String, Object> map = new HashMap();
        map.put("id", 1);
        map.put("data", 2);
        source.add(map);
        sources.add(source);

        //exportToExcel(list,arr,sources,null);
    }

    public static void setCellValue(Cell cell, Field field, Object target, Object type) throws IllegalAccessException {
        //实体类型
        Class clazz = field.getType();
        if (int.class == clazz) {
            cell.setCellValue((int) field.get(target));
        } else if (String.class == clazz) {
            cell.setCellValue((String) field.get(target));
        } else if (Date.class == clazz) {
            Date date = (Date) field.get(target);
            if (date != null)
                cell.setCellValue(dateFormat.format(date));
        }
    }

    public static void createSheetWithData(XSSFWorkbook wb, List<ExcelConfig> config, String classFullQualifiedName, List<Map<String, Object>> source, String sheetName) throws IllegalAccessException, ClassNotFoundException, NoSuchFieldException {

        XSSFSheet sheet = (XSSFSheet) createSheet(wb, config, sheetName);
        for (int i = 0; i < source.size(); i++) {
            Row row = sheet.createRow(i + 1);// 不要覆盖表头
            row.createCell(0).setCellValue((String) source.get(i).get("id"));
            for (int j = 1; j < config.size(); j++) { //0位预留给id
                Cell cell = row.createCell(j);
                XSSFCellStyle cellCellStyle = wb.createCellStyle();
                XSSFDataFormat cellDataFormat = wb.createDataFormat();
                int formatType = (config.get(j).type instanceof Integer)
                        ? (int) config.get(j).type : cellDataFormat.getFormat((String) config.get(j).type);
                cellCellStyle.setDataFormat(formatType);
                cell.setCellStyle(cellCellStyle);
                ClassLoader loader = Thread.currentThread().getContextClassLoader();
                Class clazz = loader.loadClass(classFullQualifiedName);
                Field field = clazz.getDeclaredField(config.get(j).field);
                field.setAccessible(true);
                //field.get(source.get(i).get("data"));
                setCellValue(cell, field, source.get(i).get("data"), config.get(j).type);
            }
        }
    }

    public static XSSFWorkbook exportToExcel(List<List<ExcelConfig>> configList, String[] classFullQualifiedNames, List<List<Map<String, Object>>> sources, String[] sheetNames) {
        XSSFWorkbook wb = new XSSFWorkbook();
        for (int i = 0; i < configList.size(); i++) {
            try {
                createSheetWithData(wb, configList.get(i), classFullQualifiedNames[i], sources.get(i), sheetNames == null ? null : sheetNames[0]);
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            } catch (ClassNotFoundException e) {
                e.printStackTrace();
            } catch (NoSuchFieldException e) {
                e.printStackTrace();
            }
        }
        return wb;
    }

    //Exception 作一次给用户看的转换
    public static List importFromExcel(InputStream inputStream, List<List<ExcelConfig>> configList, String[] classFullQualifiedNames) throws NoSuchFieldException, IOException, ClassNotFoundException, NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException {
        XSSFWorkbook wb = new XSSFWorkbook(inputStream);
        List result = new ArrayList();
        List<String> exceptionList = new ArrayList();
        for (int i = 0; i < configList.size(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            System.out.println("表" + i + "数据行数 :" + (sheet.getLastRowNum()));
            List<Map<String, Object>> sheetList = new ArrayList();
            ClassLoader loader = Thread.currentThread().getContextClassLoader();
            Class clazz = loader.loadClass(classFullQualifiedNames[i]);
            Constructor constructor = clazz.getDeclaredConstructor((Class[]) null);
            constructor.setAccessible(true);
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {//从1开始  忽略表头
                Row row = sheet.getRow(j);
                //略过空行
                if (row == null)
                    continue;
                String startCellVal = null;
                try {
                    //处理主键信息
                    if (row.getCell(0) == null ||
                            com.casking.chip.mdm.admin.utils.StringUtils.isBlank(
                                    startCellVal = (String) dataProcess(row.getCell(0), 49))) {
                        throw new Exception("");
                    }
                } catch (Exception e) {
                    CellReference cellRef = new CellReference(row.getRowNum(), 0);
                    exceptionList.add(sheet.getSheetName() + " 表、" + cellRef.formatAsString() + " 解析出现异常:" + e.getCause().toString());
                    continue; // 退出当前行 不会加入resultList
                }
                Map<String, Object> map = new HashMap<String, Object>();
                map.put("id", startCellVal);
                // TODO   获取单元格位置
                // CellReference cellRef = new CellReference(row.getRowNum(), row.getCell(0).getColumnIndex());
                //throw new Exception(sheet.getSheetName()+"的"+cellRef.formatAsString()+"单元格: 唯一标识(表头第一列)解析出错");
                Object obj = constructor.newInstance();
                int begin = i == 0 ? 0 : 1;//第一张表读取所有数据 其他的第一个字段作为主键索引 保存到map.data 中
                for (int k = begin; k < configList.get(i).size(); k++) {
                    Cell cell = row.getCell(k);
                    if (cell == null) continue;
                    try {
                        //空则不覆盖默认设置 如mergeflag 等
                        Object val = dataProcess(cell, configList.get(i).get(k).type);
                        if (val != null) {
                            Field field = clazz.getDeclaredField(configList.get(i).get(k).field);
                            field.setAccessible(true);
                            field.set(obj, val);
                        }
                    } catch (Exception e) {
                        CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                        exceptionList.add(sheet.getSheetName() + " 表、" + cellRef.formatAsString() + " 解析出现异常:" + e.getCause().toString());
                        //throw new Exception(sheet.getSheetName()+" 表、"+cellRef.formatAsString()+" 解析出现异常:"+e.getCause().toString());
                    } catch (Exception e) {
                        e.printStackTrace();
                        CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                        exceptionList.add(sheet.getSheetName() + " 表、" + cellRef.formatAsString() + " 解析出现异常:" + e.getCause().toString());
                        //TODO 不添加处理  直接略过该行
                        continue;
                    }
                }
                map.put("data", obj);
                sheetList.add(map);
            }
            result.add(sheetList);
        }
        result.add(exceptionList);
        return result;
    }


    public static void generateTemplate(String filepath, List<List<ExcelConfig>> list, String[] sheetName) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        for (int i = 0; i < list.size(); i++)//Sheet sheet =
            createSheet(wb, list.get(i), sheetName == null ? null : sheetName[i]);
        FileOutputStream fos = new FileOutputStream(filepath);
        wb.write(fos);
        fos.flush();
        fos.close();
    }

    public static Sheet createSheet(XSSFWorkbook wb, List<ExcelConfig> list, String name) {
        name = name == null ? "Sheet" + wb.getNumberOfSheets() : name;
        XSSFSheet sheet = wb.createSheet(name);
        sheet.setDefaultColumnWidth(12);
        Row row = sheet.createRow(0);
        for (int i = 0; i < list.size(); i++) {
            row.createCell(i).setCellValue(list.get(i).description);
            if (list.get(i).type != null) {
                //设置输入格式限定
                XSSFCellStyle xssfCellStyle = wb.createCellStyle();
                XSSFDataFormat xssfDataFormat = wb.createDataFormat();
                if (list.get(i).type instanceof Integer) {
                    xssfCellStyle.setDataFormat((int) list.get(i).type);
                    sheet.setDefaultColumnStyle(i, xssfCellStyle);
                } else {//支持原poi可扩展的format格式
                    xssfCellStyle.setDataFormat(xssfDataFormat.getFormat((String) list.get(i).type));
                    sheet.setDefaultColumnStyle(i, xssfCellStyle);
                }
            } else {
                // 默认为null 时 设置为String
                //2018.1.16 将null 在构造器优化成了 49
                System.err.println("待处理的错误(-1)");
            }
            if (list.get(i).dataValidationConstraint == null) ;
            else {
                XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
                XSSFDataValidationConstraint dvConstraint = list.get(i).dataValidationConstraint;
                CellRangeAddressList addressList = new CellRangeAddressList(1, maxRowToValidate, i, i);
                XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
                        dvConstraint, addressList);
                validation.setSuppressDropDownArrow(true);//默认值 可删吧
                if (DataValidationConstraint.ValidationType.DATE == list.get(i).dataValidationConstraint.getValidationType()) {
                    validation.createPromptBox("日期格式说明", "1900/1/1");
                    validation.setShowPromptBox(true);
                }
                validation.setShowErrorBox(true);
                sheet.addValidationData(validation);
            }
        }
        //↑ 设置列的默认格式
        //↓ 设置表头格式 //优先级好像没有列的默认格式高
        XSSFCellStyle rowCellStyle = wb.createCellStyle();
        XSSFDataFormat rowDataFormat = wb.createDataFormat();
        rowCellStyle.setDataFormat(rowDataFormat.getFormat("@"));
        row.setRowStyle(rowCellStyle);
        return sheet;
    }

}
*/
