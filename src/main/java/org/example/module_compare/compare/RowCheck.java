package org.example.module_compare.compare;

import org.apache.poi.ss.usermodel.Sheet;

public interface RowCheck {
    boolean isRowEqual(Sheet row1, Sheet row2);
}
