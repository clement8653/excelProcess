import Biz.ExcelProcessUtil;
import VO.Result;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws IOException {
//        System.out.println(ExcelProcessUtil.getSourceFilePath() + "source1.xlsx");
//        ExcelProcessUtil.exportByFormatForDrillDown(ExcelProcessUtil.getSourceFilePath() + "source1.xlsx");
        Map<String, ArrayList<Result>> map = ExcelProcessUtil.changeToMapFromExcel(ExcelProcessUtil.getSourceFilePath() + "result.xlsx");
        ExcelProcessUtil.generateResultToExcel(map);
    }
}
