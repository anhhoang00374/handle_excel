package org.example;
import org.example.module_compare.CompareFile;
import java.io.IOException;
import java.nio.file.Path;
//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) throws IOException {
        var fis = Path.of("").toAbsolutePath() + "/excel";
        String path1 = fis + "/file1.xlsx";
        String path2 = fis + "/file2.xlsx";
        CompareFile.excute(path1, path2);
    }
}