import data.ExceptionList;
import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class AlgOpen {
    public AlgOpen(InfoList infoList) throws IOException, InvalidFormatException {
        File file = new File(Application.rootDirPath + "\\algs.xlsx");
        String filePath = file.getPath();
        Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
        for(int i = 2, k = 0; i < workbook.getSheetAt(1).getPhysicalNumberOfRows();i++, k++)
        {
            if(workbook.getSheetAt(1).getRow(i).getCell(0) != null && !workbook.getSheetAt(1).getRow(i).getCell(0).getStringCellValue().equals(""))
            {
                infoList.algs.add(new ArrayList<>());
                infoList.algs.get(k).add(workbook.getSheetAt(1).getRow(i).getCell(0).getStringCellValue());
                infoList.algs.get(k).add(Double.toString(workbook.getSheetAt(1).getRow(i).getCell(3).getNumericCellValue()));
                infoList.algs.get(k).add(Double.toString(workbook.getSheetAt(1).getRow(i).getCell(4).getNumericCellValue()));
                if (workbook.getSheetAt(1).getRow(i).getCell(5) != null){
                    if(workbook.getSheetAt(1).getRow(i).getCell(5).getCellType().equals(CellType.NUMERIC)){
                        infoList.algs.get(k).add(Double.toString(workbook.getSheetAt(1).getRow(i).getCell(5).getNumericCellValue()));
                    } else {
                        if(workbook.getSheetAt(1).getRow(i).getCell(5).getStringCellValue().equals("")){
                            infoList.algs.get(k).add(workbook.getSheetAt(1).getRow(i).getCell(5).getStringCellValue());
                        } else {
                            infoList.algs.get(k).add("");
                        }
                    }
                } else {
                    infoList.algs.get(k).add("");
                }
                if(workbook.getSheetAt(1).getRow(i).getCell(7).getCellType().equals(CellType.STRING)){
                    infoList.algs.get(k).add(workbook.getSheetAt(1).getRow(i).getCell(7).getStringCellValue());
                } else {
                    infoList.algs.get(k).add(Double.toString(workbook.getSheetAt(1).getRow(i).getCell(7).getNumericCellValue()));
                }
                if(workbook.getSheetAt(1).getRow(i).getCell(8).getCellType().equals(CellType.STRING)){
                    infoList.algs.get(k).add(workbook.getSheetAt(1).getRow(i).getCell(8).getStringCellValue());
                } else {
                    infoList.algs.get(k).add(Double.toString(workbook.getSheetAt(1).getRow(i).getCell(8).getNumericCellValue()));
                }
                if(workbook.getSheetAt(1).getRow(i).getCell(9).getCellType().equals(CellType.STRING)){
                    infoList.algs.get(k).add(workbook.getSheetAt(1).getRow(i).getCell(9).getStringCellValue());
                } else {
                    infoList.algs.get(k).add(Double.toString(workbook.getSheetAt(1).getRow(i).getCell(9).getNumericCellValue()));
                }
                if(workbook.getSheetAt(1).getRow(i).getCell(10).getCellType().equals(CellType.STRING)){
                    infoList.algs.get(k).add(workbook.getSheetAt(1).getRow(i).getCell(10).getStringCellValue());
                    if(!infoList.uniqBact.contains(workbook.getSheetAt(1).getRow(i).getCell(10).getStringCellValue())){
                        infoList.uniqBact.add(workbook.getSheetAt(1).getRow(i).getCell(10).getStringCellValue());
                    }
                } else {
                    infoList.algs.get(k).add(Double.toString(workbook.getSheetAt(1).getRow(i).getCell(10).getNumericCellValue()));
                }
            }
        }
        workbook.close();
    }
}
