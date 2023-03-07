package fileView;

import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class XLXSOpen {
    String fileName;
    Workbook workbook;
    public XLXSOpen(File file) throws IOException, InvalidFormatException {
        String filePath = file.getPath();
        fileName = file.getName();
        workbook = new XSSFWorkbook(new FileInputStream(filePath));
    }

    public void getBioIndex(InfoList infoList){
        boolean checkAvailability = false;
        for(int i = 0; i < workbook.getNumberOfSheets(); i++){
            if (workbook.getSheetAt(i).getSheetName().contains("BioIndex")){
                checkAvailability = true;
                break;
            }
        }
        if(checkAvailability){
            String checkIndex = String.format("%(.2f", Double.parseDouble(workbook.getSheet("BioIndex").getRow(0).getCell(1).getStringCellValue()));
            infoList.bioIndex.add(checkIndex);
            infoList.bioIndex.add(workbook.getSheet("BioIndex").getRow(0).getCell(2).getStringCellValue());
        }
    }

    public void getPielouEveness(InfoList infoList){
        boolean checkAvailability = false;
        for(int i = 0; i < workbook.getNumberOfSheets(); i++){
            if (workbook.getSheetAt(i).getSheetName().contains("Eveness")){
                checkAvailability = true;
                break;
            }
        }
        if(checkAvailability){
            infoList.pielouEveness = String.format("%(.2f", workbook.getSheet("Eveness").getRow(0).getCell(1).getNumericCellValue());
        }
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void getPhylum(InfoList infoList) throws IOException {
        boolean checkAvailability = false;
        for(int i = 0; i < workbook.getNumberOfSheets(); i++){
            if (workbook.getSheetAt(i).getSheetName().contains("Phylum")){
                checkAvailability = true;
                break;
            }
        }
        if(checkAvailability) {
            int temp = 0;
            for (int i = 0; i < workbook.getSheet("Phylum").getPhysicalNumberOfRows(); i++) {
                if (!workbook.getSheet("Phylum").getRow(i).getCell(0).getStringCellValue().equals("Bacterium name")) {
                    infoList.phylum.add(new ArrayList<>());
                    infoList.phylum.get(i - temp).add(workbook.getSheet("Phylum").getRow(i).getCell(0).getStringCellValue());
                    if (workbook.getSheet("Phylum").getRow(i).getCell(1).getCellType().equals(CellType.NUMERIC)) {
                        Double num = workbook.getSheet("Phylum").getRow(i).getCell(1).getNumericCellValue();
                        infoList.phylum.get(i - temp).add(Double.toString(num));
                    } else {
                        String num = workbook.getSheet("Phylum").getRow(i).getCell(1).getStringCellValue();
                        infoList.phylum.get(i - temp).add(num);
                    }
                } else {
                    temp++;
                }
            }
        }
    }

    public void getGenus(InfoList infoList) throws IOException {
        int temp = 0;
        for(int i = 0; i < workbook.getSheet("Genus").getPhysicalNumberOfRows();i++)
        {
            if (!workbook.getSheet("Genus").getRow(i).getCell(0).getStringCellValue().equals("Bacterium name")){
                infoList.genus.add(new ArrayList<>());
                infoList.genus.get(i-temp).add(workbook.getSheet("Genus").getRow(i).getCell(0).getStringCellValue());
                if(workbook.getSheet("Genus").getRow(i).getCell(1).getCellType().equals(CellType.NUMERIC)){
                    Double num = workbook.getSheet("Genus").getRow(i).getCell(1).getNumericCellValue();
                    infoList.genus.get(i-temp).add(Double.toString(num));
                } else {
                    String num = workbook.getSheet("Genus").getRow(i).getCell(1).getStringCellValue();
                    infoList.genus.get(i-temp).add(num);
                }
            } else {
                temp++;
            }
        }
    }

    public void getSpecies(InfoList infoList) throws IOException {
        int temp = 0;
        for(int i = 0; i < workbook.getSheet("Species").getPhysicalNumberOfRows();i++)
        {
            if (!workbook.getSheet("Species").getRow(i).getCell(0).getStringCellValue().equals("Bacterium name")){
                infoList.species.add(new ArrayList<>());
                infoList.species.get(i-temp).add(workbook.getSheet("Species").getRow(i).getCell(0).getStringCellValue());
                if(workbook.getSheet("Species").getRow(i).getCell(1).getCellType().equals(CellType.NUMERIC)){
                    Double num = workbook.getSheet("Species").getRow(i).getCell(1).getNumericCellValue();
                    infoList.species.get(i-temp).add(Double.toString(num));
                } else {
                    String num = workbook.getSheet("Species").getRow(i).getCell(1).getStringCellValue();
                    infoList.species.get(i-temp).add(num);
                }
            } else {
                temp++;
            }
        }
    }

    public void getFamily(InfoList infoList) throws IOException {
        int temp = 0;
        for(int i = 0; i < workbook.getSheet("Family").getPhysicalNumberOfRows();i++)
        {
            if (!workbook.getSheet("Family").getRow(i).getCell(0).getStringCellValue().equals("Bacterium name")){
                infoList.family.add(new ArrayList<>());
                infoList.family.get(i-temp).add(workbook.getSheet("Family").getRow(i).getCell(0).getStringCellValue());
                if(workbook.getSheet("Family").getRow(i).getCell(1).getCellType().equals(CellType.NUMERIC)){
                    Double num = workbook.getSheet("Family").getRow(i).getCell(1).getNumericCellValue();
                    infoList.family.get(i-temp).add(Double.toString(num));
                } else {
                    String num = workbook.getSheet("Family").getRow(i).getCell(1).getStringCellValue();
                    infoList.family.get(i-temp).add(num);
                }
            } else {
                temp++;
            }
        }
    }

    public void getFileName(InfoList infoList){
        infoList.fileName = fileName;
    }
}