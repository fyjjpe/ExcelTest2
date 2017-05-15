package excel;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;

/**
 * POI解析excel文件
 * Created by Administrator on 2017/5/16.
 */
public class PoiReadExcel {
    public static void main(String[] args) {
        //需要解析的文件
        File file = new File("d:/poi_test.xls");
        try {
            //创建excel,读取文件内容
            HSSFWorkbook workbook = new HSSFWorkbook(FileUtils.openInputStream(file));

            //获取Sheet0工作表
//            HSSFSheet sheet = workbook.getSheet("Sheet0")；
            //获取Sheeto工作表的另一种方式
            HSSFSheet sheet = workbook.getSheetAt(0);
            int firstRowNum = 0;
            //获取最后一行行号
            int lastRowNum = sheet.getLastRowNum();
            for (int i = firstRowNum; i <= lastRowNum; i++) {
                HSSFRow row = sheet.getRow(i);
                //获取当前行中最后一列
                int lastCellNum = row.getLastCellNum();
                for (int j = 0; j < lastCellNum ; j++) {
                    HSSFCell cell = row.getCell(j);
                    String value = cell.getStringCellValue();
                    System.out.print(value + "  ");
                }
                System.out.println();
            }


        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
