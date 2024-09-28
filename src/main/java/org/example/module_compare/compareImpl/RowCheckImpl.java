package org.example.module_compare.compareImpl;

import org.apache.poi.ss.usermodel.Sheet;
import org.example.module_compare.compare.RowCheck;

public class RowCheckImpl implements RowCheck {
    @Override
    public boolean isRowEqual(Sheet sheet1, Sheet sheet2) {
        var check = true;
        if(sheet1.getLastRowNum() != sheet2.getLastRowNum()){
            System.out.println("ERROR - Số hàng của 2 file đang kháo nhau");
            return false;
        }
        for (int i = 0; i < sheet1.getLastRowNum(); i++) {
            if(sheet1.getRow(i).getHeight() != sheet2.getRow(i).getHeight()){
                System.out.println("ERROR - Row height chiều cao của 2 hàng khác nhau ở:" + (i+1));
                check = false;
            }
        }
        return check;
    }
}
