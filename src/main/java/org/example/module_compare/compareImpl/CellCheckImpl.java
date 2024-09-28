package org.example.module_compare.compareImpl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.example.module_compare.compare.CellCheck;

import java.util.List;

public class CellCheckImpl implements CellCheck {
    String[] code = {
            "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
            "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
            "U", "V", "W", "X", "Y", "Z",
            "AA", "AB", "AC", "AD"
    };
    @Override
    public boolean checkCell(Sheet sheet1, Sheet sheet2, Workbook workbook,Workbook workbook1) {
        var check = true;
        for(int index = 0; index < sheet1.getLastRowNum(); index++){
            Row rowSheet1 = sheet1.getRow(index);
            Row rowSheet2 = sheet2.getRow(index);
            for(int j = 0; j < rowSheet1.getLastCellNum(); j++){
                Cell cell1 = rowSheet1.getCell(j);
                Cell cell2 = rowSheet2.getCell(j);
                if((cell1 != null && cell2 == null)){
                    System.out.println("ERROR - Empty " + code[cell1.getColumnIndex()]+ cell1.getRowIndex() + " đang có giá trị rỗng");
                    check = false;
                }
                if((cell1 == null && cell2 != null)){
                    System.out.println("ERROR - Empty " + code[cell2.getColumnIndex()]+ cell2.getRowIndex() + " đang có giá trị rỗng");
                    check = false;
                }
                if(cell1 != null && cell2 != null){
                    //type
                    if(!checkType(cell1,cell2)){
                        check = false;
                        System.out.println("ERROR - Type " + code[cell2.getColumnIndex()]+ (cell2.getRowIndex() + 1) + " đang có kiểu dữ liệu khác nhau");
                    }
                    //value
                    if(!checkValue(cell1,cell2)){
                        check = false;
                        System.out.println("ERROR - Value" + code[cell2.getColumnIndex()]+ (cell2.getRowIndex() + 1) + " đang có giá trị khác nhau");
                    }

                    //check style bold, italic, underline, font size, font name, color
                    if(!checkStyle(cell1,cell2,workbook,workbook1)){
                        check = false;
                    }

                    //check border
                    if(!checkBorder(cell1,cell2)){
                        check = false;
                    }

                    //check border
                    if(!checkBackground(cell1,cell2)){
                        System.out.println("ERROR - Background " + code[cell2.getColumnIndex()]+ (cell2.getRowIndex() + 1) + " đang có giá trị khác nhau");
                        check = false;
                    }

                    //check border
                    if(!checkAlignVerticel(cell1,cell2)){
                        check = false;
                    }
                }
            }

        }
        return check;
    }

    private boolean checkAlignVerticel(Cell cell1, Cell cell2) {
        var check = true;
        CellStyle style1 = cell1.getCellStyle();
        CellStyle style2 = cell2.getCellStyle();
        if(!(style1.getAlignment() == style2.getAlignment())){
            check = false;
            System.out.println("ERROR - Align " + code[cell2.getColumnIndex()]+ (cell2.getRowIndex() + 1) + " đang có giá trị khác nhau");
        }

        if(!(style1.getVerticalAlignment() == style2.getVerticalAlignment())){
            check = false;
            System.out.println("ERROR - Vertical " + code[cell2.getColumnIndex()]+ (cell2.getRowIndex() + 1) + " đang có giá trị khác nhau");
        }

        return check;
    }

    private boolean checkBackground(Cell cell1, Cell cell2){
        CellStyle style1 = cell1.getCellStyle();
        CellStyle style2 = cell2.getCellStyle();
        return style1.getFillForegroundColor() == style2.getFillForegroundColor();
    }

    private boolean checkBorder(Cell cell1, Cell cell2){
        var check = true;
        CellStyle style1 = cell1.getCellStyle();
        CellStyle style2 = cell2.getCellStyle();
        if(!style1.getBorderBottom().name().equals(style2.getBorderBottom().name())){
            System.out.println("ERROR - Border bottom của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            check = false;
        }

        if(!style1.getBorderTop().name().equals(style2.getBorderTop().name())){
            System.out.println("ERROR - Border top của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            check = false;
        }

        if(!style1.getBorderLeft().name().equals(style2.getBorderLeft().name())){
            System.out.println("ERROR - Border left của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            check = false;
        }

        if(!style1.getBorderRight().name().equals(style2.getBorderRight().name())){
            System.out.println("ERROR - Border right của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            check = false;
        }
        return check;
    }

    private boolean checkStyle(Cell cell1, Cell cell2,Workbook workbook,Workbook workbook2){
        var check = true;
        if(checkValue(cell1,cell2)){
            CellStyle style = cell1.getCellStyle();
            Font font = workbook.getFontAt(style.getFontIndex());
            CellStyle style2 = cell2.getCellStyle();
            Font font2 = workbook2.getFontAt(style2.getFontIndex());
            if(!(font.getBold() == font2.getBold())){
                check = false;
                System.out.println("ERROR - Bold của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            }
            if(!(font.getItalic() == font2.getItalic())){
                check = false;
                System.out.println("ERROR - Italic của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            }

            if(!(font.getUnderline() == font2.getUnderline())){
                check = false;
                System.out.println("ERROR - Underline của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            }

            if(!(font.getFontName().equals(font2.getFontName()))){
                check = false;
                System.out.println("ERROR - Font chữ của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            }

            if(!(font.getFontHeight() == font2.getFontHeight())){
                check = false;
                System.out.println("ERROR - Font size của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            }

            if(!(font.getColor() == font2.getColor())){
                check = false;
                System.out.println("ERROR - Color màu chữ của " + code[cell1.getColumnIndex()]+ (cell1.getRowIndex() + 1) + " đang có giá trị khác nhau");
            }

        }
        return check;

    }
    private boolean checkType(Cell cell1, Cell cell2){
        return cell1.getCellType() == cell2.getCellType();
    }

    private boolean checkValue(Cell cell1, Cell cell2){
        if(cell1.getCellType() == cell2.getCellType()){
            switch (cell1.getCellType()) {
                case STRING:
                    return cell1.getStringCellValue().trim().equals(cell2.getStringCellValue());
                case NUMERIC:
                    return cell1.getNumericCellValue() == cell2.getNumericCellValue();
                case BOOLEAN:
                    return cell1.getBooleanCellValue() == cell2.getBooleanCellValue();
                case FORMULA:
                    return cell1.getCellFormula().equals(cell2.getCellFormula());
                case BLANK:
                    return cell1.getStringCellValue().equals(cell2.getStringCellValue());
                default:
                    return false;
            }
        }
        return false;
    }

}
