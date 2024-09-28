package org.example.module_compare.compareImpl;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.example.module_compare.compare.MergedCheck;

import java.util.List;

public class MergedCheckImpl implements MergedCheck {
    String[] code = {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
            "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
            "U", "V", "W", "X", "Y", "Z",
            "AA", "AB", "AC", "AD"
    };
    @Override
    public boolean checkMerged(Sheet sheet1, Sheet sheet2) {
        var check = true;
        List<CellRangeAddress> mergerRange = sheet1.getMergedRegions();
        List<CellRangeAddress> mergerRange2 = sheet2.getMergedRegions();
        if(mergerRange.size() != mergerRange2.size()){
            System.out.println("ERROR - Merge Số lượng ô được merge của hai file đang khác nhau :" + mergerRange.size() + "và" + mergerRange2.size());
            var number = Math.max(mergerRange.size(), mergerRange2.size());
            List<CellRangeAddress> checkSheet = mergerRange.size() > mergerRange2.size() ? mergerRange : mergerRange2;
            for(int i = 0; i < number; i++){
                CellRangeAddress cellRangeAddress = checkSheet.get(i);
                System.out.println("ERROR - Merge sheet được merged ở vị trí hàng " + cellRangeAddress.getFirstRow() + " và cột " + code[cellRangeAddress.getLastColumn()]);
            }
            check = false;
            return check;
        }

        var checkMerge = true;
        var number = Math.min(mergerRange.size(), mergerRange2.size());
        for(int i = 0 ; i < number ; i++){
            if(!compareMerged(sheet1.getMergedRegion(i),sheet2.getMergedRegion(i))){
                checkMerge = false;
            }
        }
        return checkMerge;
    }

    private boolean compareMerged(CellRangeAddress cellRangeAddress1, CellRangeAddress cellRangeAddress2) {
        var check = true;
        if(cellRangeAddress1.getFirstRow() != cellRangeAddress2.getFirstRow()){
            System.out.println("ERROR - Merge ô được merge khác nhau ở hàng :" + (cellRangeAddress1.getFirstRow() + 1) + " và " + (cellRangeAddress1.getFirstRow() + 1));
            check = false;
        }
        if(cellRangeAddress1.getLastRow() != cellRangeAddress2.getLastRow()){
            System.out.println("ERROR - Merge ô được merge khác nhau ở hàng :" + (cellRangeAddress1.getLastRow() + 1) + " và " + (cellRangeAddress1.getLastRow() + 1));
            check = false;
        }
        if(cellRangeAddress1.getFirstColumn() != cellRangeAddress2.getFirstColumn()){
            System.out.println("ERROR - Merge ô được merge khác nhau ở cột :" + code[cellRangeAddress1.getFirstColumn()] + " và " + code[cellRangeAddress2.getFirstColumn()]);
            check = false;
        }
        if(cellRangeAddress1.getLastColumn() != cellRangeAddress2.getLastColumn()){
            System.out.println("ERROR - Merge ô được merge khác nhau ở cột :" + code[cellRangeAddress1.getLastColumn()] + " và " + code[cellRangeAddress2.getLastColumn()]);
            check = false;
        }
        return check;
    }
}
