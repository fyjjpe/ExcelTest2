package excel;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

/**
 * 使用POI生成excel文件
 * Created by Administrator on 2017/5/16.
 */
public class PoiExpExcel {
    public static void main(String[] args) {
        String[] title = {"id", "name", "sex"};
        //创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建工作表
        HSSFSheet sheet = workbook.createSheet();
        //创建第一行
        HSSFRow row = sheet.createRow(0);

        HSSFCell cell = null;
        //插入第一行数据
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }

        //追加插入
        for (int i = 1; i <= 10; i++) {
            HSSFRow newRow = sheet.createRow(i);
            HSSFCell cell1 = newRow.createCell(0);
            cell1.setCellValue("A" + i);
            cell1 = newRow.createCell(1);
            cell1.setCellValue("user" + i);
            cell1 = newRow.createCell(2);
            cell1.setCellValue("男");
        }

        //创建保存的文件
        File file = new File("d:/poi_test.xls");
        try {
            file.createNewFile();
            FileOutputStream stream = FileUtils.openOutputStream(file);
            workbook.write(stream);
            workbook.close();
        }catch (Exception e){
            e.printStackTrace();
        }

    }
}

