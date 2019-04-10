package trev0r;

/**
 * @author： huanggq
 * @CreateTime: 2019/4/11
 * @lastModTime: 2019/4/11
 * @version: 1.0
 * @description:
 */
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;

public class ExcelConfig {
    public static final String sep = "\\";
    public static final String filePrefix = "chip-mdm-parent" + sep + "chip-mdm-admin" + sep + "src" + sep + "main" + sep + "resources" + sep + "temp" + sep;
    public static final String fileSuffix = ".xlsx";

    public static String getFilePath(String fileName) {
        return filePrefix + fileName + fileSuffix;
    }
    /**
     * type 按需做formatter限定
     * null: 不限定
     *      int:配置在下面类中
     *          org.apache.poi.ss.usermodel.BuiltinFormats
     *          其中 14 日期
     *              49 文字
     * string: 自定格式规则
     */

    public String field;
    public String description;
    public Object type;
    public XSSFDataValidationConstraint dataValidationConstraint;
    /**
     * @description 可选配置在DataValidationConstraint中
     * @option1  添加枚举选项
     *      new XSSFDataValidationConstraint(new String[]{"1","2","2"})
     * @option2  限制长度 arg0 长度限制 arg1 作用类型(between less than 等) arg2 arg3和arg1 相关
     *      new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.TEXT_LENGTH, 0x05,"2","")
     * @option3  限制日期
     *      new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.DATE, DataValidationConstraint.OperatorType.GREATER_THAN,"1900/1/1","")
     */

    public ExcelConfig(String field,String description,Object type,XSSFDataValidationConstraint dataValidationConstraint){
        this.field = field;
        this.description = description;
        this.type = type==null ? 49 :type;
        this.dataValidationConstraint = dataValidationConstraint;
    }

    public ExcelConfig(String field,String description,Object type){
        this(field,description,type,null);
    }
    public ExcelConfig(String field,String description){
        this(field,description,49,null);
    }

    public static XSSFDataValidationConstraint getLessthanLengthConstraint(String length) {
        return new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.TEXT_LENGTH, DataValidationConstraint.OperatorType.LESS_THAN,length,"");
    }

    public String getField() {
        return field;
    }

    public void setField(String field) {
        this.field = field;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public Object getType() {
        return type;
    }

    public void setType(Object type) {
        this.type = type;
    }
}




