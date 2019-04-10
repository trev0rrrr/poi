package trev0r;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;

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
        XSSFSheet sheet = workbook.createSheet("sheet2");
        System.out.println(1_000_000);
        for (int i = 0; i < 1; i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(0);
/*            XSSFCellStyle cellStyle = new XSSFCellStyle(new XSSFBStylesTable());
            cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
            cell.setCellStyle(cellStyle);*/
            cell.setCellValue("hello");
            Cell cell2 = row.createCell(1);
            cell2.setCellValue(1.2d);
            Cell cell3 = row.createCell(2);
            cell3.setCellValue(new Date());
            Cell cell4 = row.createCell(3);
            cell4.setCellValue(true);
        }

        workbook.write(new FileOutputStream(f));
        Process p = Runtime.getRuntime().exec("cmd /c start " + f.getAbsolutePath());
        while(p.isAlive()){

        }
        System.out.println("cmd /c start " + f.getAbsolutePath());

    }

    public static void genTemplate(List list,String[] sheetName) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        for (int i = 0; i < list.size(); i++)//Sheet sheet =
            createSheet(wb, list.get(i), sheetName == null ? null : sheetName[i]);
        FileOutputStream fos = new FileOutputStream("workbook.xlsx");
        wb.write(fos);
        fos.flush();
        fos.close();
    }

    public static void createSheet(XSSFWorkbook wb,Object o ,String name) {
        name = name == null ? "Sheet" + wb.getNumberOfSheets() : name;
        XSSFSheet sheet = wb.createSheet(name);
//        sheet.setDefaultColumnWidth(12);
        Row row = sheet.createRow(0);

    }
}
