package org.example.module_compare.compare;
import org.apache.poi.ss.usermodel.Sheet;

public interface ColumnCheck {
    boolean isColumnEqual(Sheet sheet1, Sheet sheet2);
}
