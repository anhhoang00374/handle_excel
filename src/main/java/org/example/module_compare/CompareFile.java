package org.example.module_compare;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.module_compare.compareImpl.CellCheckImpl;
import org.example.module_compare.compareImpl.ColumnCheckImpl;
import org.example.module_compare.compareImpl.MergedCheckImpl;
import org.example.module_compare.compareImpl.RowCheckImpl;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class CompareFile {
    public static boolean excute(String path1, String path2) throws IOException {
        var check = false;
        Workbook workbook =  getWorkbook(path1);
        Workbook workbook2 = new XSSFWorkbook(path2);
        Sheet sheet = workbook.getSheetAt(0);
        Sheet sheet2 = workbook2.getSheetAt(0);

        //kiểm tra độ rộng cột
        ColumnCheckImpl columnCheck = new ColumnCheckImpl();
        if(columnCheck.isColumnEqual(sheet, sheet2)){
            System.out.println("OK - Độ rộng cột của 2 sheet bằng nhau");
            check = true;
        }

        //kiểm tra chiều cao của hàng
        RowCheckImpl rowCheck = new RowCheckImpl();
        if(rowCheck.isRowEqual(sheet, sheet2)){
            System.out.println("OK - Chiều cao hàng của 2 sheet bằng nhau");
            check = true;
        }

        //kiểm tra các ô được merge
        MergedCheckImpl mergedCheck = new MergedCheckImpl();
        if(mergedCheck.checkMerged(sheet, sheet2)){
            System.out.println("OK - Các ô được merge 2 file giống nhau");
            check = true;
        }
        if(!check){
            return check;
        }

        //kiểm tra từng ô
        CellCheckImpl cellCheck = new CellCheckImpl();
        if(cellCheck.checkCell(sheet, sheet2,workbook,workbook2)){
            System.out.println("OK - Tất cả các ô đã giống nhau");
        }
        return check;
    }
    private static Workbook getWorkbook(String path) throws IOException {
        File f1 = new File(path);
        FileInputStream fInput1 = new FileInputStream(f1);
        Workbook workbook = new XSSFWorkbook(fInput1);
        return workbook;
    }
}
