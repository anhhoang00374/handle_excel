package org.example.module_compare.compareImpl;

import org.apache.poi.ss.usermodel.Sheet;
import org.example.module_compare.compare.ColumnCheck;

public class ColumnCheckImpl implements ColumnCheck {

    @Override
    public boolean isColumnEqual(Sheet sheet1, Sheet sheet2) {
        String[] code = {
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
                "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
                "U", "V", "W", "X", "Y", "Z",
                "AA", "AB", "AC", "AD"
        };
        var check = true;
        int lastColumn = 30;
        for (int i = 0; i < lastColumn; i++) {
            if(sheet1.getColumnWidth(i) != sheet2.getColumnWidth(i)) {
                System.out.println("ERROR - Column width độ rộng của cột khác nhau ở cột: " + code[i]);
                check = false;
            }
        }
        return check;
    }
}
