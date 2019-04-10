package trev0r;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

/**
 * Hello world!
 */
public class App {
    public static void init(String[] args) {

    }

    public static void main(String[] args) throws Exception {
        String file = "workbook.xlsx";
        File f = new File(file);
        if (f.exists())
            f.delete();
        XSSFWorkbook workbook = new XSSFWorkbook();
//        workbook.createSheet("sheet1");
        workbook.createSheet("sheet2");
        workbook.write(new FileOutputStream(f));
        Process p = Runtime.getRuntime().exec("cmd /c start " + f.getAbsolutePath());
        while(p.isAlive()){

        }
        System.out.println("cmd /c start " + f.getAbsolutePath());

    }
}
