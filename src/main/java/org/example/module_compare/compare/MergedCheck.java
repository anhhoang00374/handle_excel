package org.example.module_compare.compare;

import org.apache.poi.ss.usermodel.Sheet;

public interface MergedCheck {
    boolean checkMerged(Sheet sheet1, Sheet sheet2);
}
