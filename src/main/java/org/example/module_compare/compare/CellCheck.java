package org.example.module_compare.compare;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface CellCheck {
    boolean checkCell(Sheet sheet1, Sheet sheet2, Workbook workbook, Workbook workbook2);
}
