import data.ExceptionList;
import data.InfoList;
import javafx.application.Platform;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.swing.*;
import java.io.*;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;

public class LoaderForObrFirst extends JFrame {
    XWPFDocument workbook;
    XWPFDocument workbookTemp;
    SortedTable sortedTable;
    String nameObr;
    public LoaderForObrFirst(String nameObr) throws IOException, InvalidFormatException {
        File file = new File(Application.rootDirPath + "\\" + nameObr + ".docx");
        workbook = new XWPFDocument(new FileInputStream(file));
        this.nameObr = nameObr;
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void setFileNameForSecond(InfoList infoList){
        XWPFRun run = workbook.getTables().get(0).getRow(0).getCell(1).getParagraphs().get(3).createRun();
        run.setFontSize(10);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.fileName.replace(".xlsx", ""));
    }

    public void setBioIndex(InfoList infoList, int numberTable){
        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(0).getCell(2).getParagraphs().get(1).createRun();
        run.setFontSize(9);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.bioIndex.get(0));
        run.setColor("0db3b3");
        run.setBold(true);
        run = workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRow(0).getCell(3).getParagraphs().get(0).createRun();
        run.setFontSize(9);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.bioIndex.get(1));
        run.setColor("0db3b3");
        run.setBold(true);
    }

    public void setDataInFiveColumnTable(InfoList infoList, int numberTable) throws XmlException, IOException {
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        for(int i = 1; i < workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRows().size();i++) {
            if (workbook.getTables().get(0).getRow(1).getCell(0).getTables().get(numberTable).getRows().get(i).getTableCells().size() == 5) {
                for (int j = 0; j < genusSpecies.size(); j++) {
                    if (genusSpecies.get(j).get(0).equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                            .get(numberTable).getRow(i).getCell(0).getText())
                            || genusSpecies.get(j).get(0).replace("_", " ").equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                            .get(numberTable).getRow(i).getCell(0).getText())
                    ) {
                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                .get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");

                        if (Double.parseDouble(genusSpecies.get(j).get(1)) == 0) {
                            run.setText(genusSpecies.get(j).get(1));
                        } else {
                            setValueGenusSpecies(run, genusSpecies, j);
                        }

                        for (int k = 0; k < infoList.algsUrogenital.size(); k++) {
                            if (infoList.algsUrogenital.get(k).get(0).equals(genusSpecies.get(j).get(0))
                                    || (infoList.algsUrogenital.get(k).get(0).contains(genusSpecies.get(j).get(0))
                                    && infoList.algsUrogenital.get(k).get(0).contains("/")
                                    && !infoList.algsUrogenital.get(k).get(0).contains(genusSpecies.get(j).get(0) + "_"))
                                    || infoList.algsUrogenital.get(k).get(0).equals(genusSpecies.get(j).get(0).
                                    replace(" ", "_"))
                                    || genusSpecies.get(j).get(0).equals(infoList.algsUrogenital.get(k).get(0).
                                    replace("_", " "))) {
                                if (checkValueRange(infoList.algsUrogenital.get(k).get(1), infoList.algsUrogenital.get(k).get(2), genusSpecies.get(j).get(1)).equals("null")) {
                                    run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                            .get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Значение отсутствует");
                                    break;
                                } else if (checkValueRange(infoList.algsUrogenital.get(k).get(1), infoList.algsUrogenital.get(k).get(2), genusSpecies.get(j).get(1)).equals("low")) {
                                    run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                            .get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Низкое значение");
                                    break;
                                } else if (checkValueRange(infoList.algsUrogenital.get(k).get(1), infoList.algsUrogenital.get(k).get(2), genusSpecies.get(j).get(1)).equals("medium")) {
                                    run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                            .get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Среднее значение");
                                    break;
                                } else if (checkValueRange(infoList.algsUrogenital.get(k).get(1), infoList.algsUrogenital.get(k).get(2), genusSpecies.get(j).get(1)).equals("high")) {
                                    run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                            .get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Высокое значение");
                                    break;
                                }
                            }
                        }
                        break;
                    } else if (j == genusSpecies.size() - 1
                            &&  !workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                            .get(numberTable).getRow(i).getCell(3).getText().equals("")
                            &&  workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                            .get(numberTable).getRow(i).getCell(4).getText().equals("")
                            &&  !workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                            .get(numberTable).getRow(i).getCell(0).getText().equals("")){
                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                .get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("Интерпретация не определена");
                        run.setColor("a60000");
                    } else if (j == genusSpecies.size() - 1
                            &&  workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                            .get(numberTable).getRow(i).getCell(3).getText().equals("")
                            &&  !workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                            .get(numberTable).getRow(i).getCell(0).getText().equals("")){
                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                .get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("0.0");
                        run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                .get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("Значение отсутствует");
                    }
                }

                for (int f = 0; f < workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                        .get(numberTable).getRow(1).getTableCells().size();f++){
                    if(workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                            .get(numberTable).getRow(1).getCell(f).getText().contains("Среднее")){
                        for(int k = 0; k<infoList.algsUrogenital.size();k++) {
                            if(!workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                    .get(numberTable).getRow(i).getCell(0).getText().equals("")){
                                if (infoList.algsUrogenital.get(k).get(0).equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                        .get(numberTable).getRow(i).getCell(0).getText())
                                        || (infoList.algsUrogenital.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                        .get(numberTable).getRow(i).getCell(0).getText())
                                        && infoList.algsUrogenital.get(k).get(0).contains("/")
                                        && !infoList.algsUrogenital.get(k).get(0).contains(workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                        .get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                        || infoList.algsUrogenital.get(k).get(0).equals(workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                        .get(numberTable).getRow(i).getCell(0).getText().
                                        replace(" ", "_"))
                                        || workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                        .get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algsUrogenital.get(k).get(0).
                                                replace("_", " "))) {
                                    if (Double.parseDouble(infoList.algsUrogenital.get(k).get(1)) == 0
                                            && Double.parseDouble(infoList.algsUrogenital.get(k).get(2)) == 0) {
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText("0-0");
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algsUrogenital.get(k).get(2)) <= 0.0009) {
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.6f", Double.parseDouble(infoList.algsUrogenital.get(k).get(1)))
                                                + "-" + String.format("%(.4f", Double.parseDouble(infoList.algsUrogenital.get(k).get(2))));
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algsUrogenital.get(k).get(1)) <= 0.0009
                                            && Double.parseDouble(infoList.algsUrogenital.get(k).get(2)) <= 0.009) {
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.4f", Double.parseDouble(infoList.algsUrogenital.get(k).get(1)))
                                                + "-" + String.format("%(.3f", Double.parseDouble(infoList.algsUrogenital.get(k).get(2))));
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algsUrogenital.get(k).get(1)) <= 0.0009
                                            && Double.parseDouble(infoList.algsUrogenital.get(k).get(2)) <= 0.09) {
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.4f", Double.parseDouble(infoList.algsUrogenital.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algsUrogenital.get(k).get(2))));
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algsUrogenital.get(k).get(1)) <= 0.009
                                            && Double.parseDouble(infoList.algsUrogenital.get(k).get(2)) <= 0.09) {
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.3f", Double.parseDouble(infoList.algsUrogenital.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algsUrogenital.get(k).get(2))));
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algsUrogenital.get(k).get(1)) <= 0.0009) {
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.4f", Double.parseDouble(infoList.algsUrogenital.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algsUrogenital.get(k).get(2))));
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algsUrogenital.get(k).get(1)) <= 0.009) {
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.3f", Double.parseDouble(infoList.algsUrogenital.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algsUrogenital.get(k).get(2))));
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else {
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.2f", Double.parseDouble(infoList.algsUrogenital.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algsUrogenital.get(k).get(2))));
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(0).getRow(1).getCell(0).getTables()
                                                .get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    }
                                }
                            }
                        }
                        break;
                    }
                }
            }
        }
    }

    public void setValueGenusSpecies(XWPFRun run, ArrayList<ArrayList<String>> genusSpecies,int j){
        if(Double.parseDouble(genusSpecies.get(j).get(1)) <= 0.0009){
            run.setText(String.format("%(.4f", Double.parseDouble(genusSpecies.get(j).get(1))));
        } else if(Double.parseDouble(genusSpecies.get(j).get(1)) <= 0.009){
            run.setText(String.format("%(.3f", Double.parseDouble(genusSpecies.get(j).get(1))));
        } else {
            run.setText(String.format("%(.2f", Double.parseDouble(genusSpecies.get(j).get(1))));
        }
    }


    String checkValueRange(String firstMediumValue, String secondMediumValue, String currentValue){
        if (Double.parseDouble(currentValue) == 0){
            return "null";
        } else if(Double.parseDouble(currentValue) >= Double.parseDouble(firstMediumValue) && Double.parseDouble(currentValue) <= Double.parseDouble(secondMediumValue)){
            return "medium";
        } else if (Double.parseDouble(currentValue) < Double.parseDouble(firstMediumValue) && Double.parseDouble(currentValue) > 0){
            return "low";
        } else if (Double.parseDouble(currentValue) > Double.parseDouble(secondMediumValue)){
            return "high";
        }
        return "null";
    }

    public void setTwoFormatWithSer(InfoList infoList, int numberTable, String choiceTable) throws IOException, ClassNotFoundException {
        ArrayList<ArrayList<String>> result = new ArrayList<>();
        if(choiceTable.equals("genus"))
        {
            result.addAll(infoList.genus);
        } else if (choiceTable.equals("species")){
            result.addAll(infoList.species);
        } else if (choiceTable.equals("family")) {
            result.addAll(infoList.family);
        }
        Collections.sort(result, new Comparator<ArrayList<String>>() {
            @Override
            public int compare(ArrayList<String> o1, ArrayList<String> o2) {
                return Double.compare(Double.parseDouble(o2.get(1)), Double.parseDouble(o1.get(1)));
            }
        });
        for(int i = 0; i < result.size();i++){
            workbook.getTables().get(numberTable).createRow();

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1)
                    .setWidth(String.valueOf(workbook.getTables().get(numberTable).getRow(0).getCell(0).getWidth()));
            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1)
                    .setWidthType(workbook.getTables().get(numberTable).getRow(0).getCell(1).getWidthType());

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0)
                    .setWidth(String.valueOf(workbook.getTables().get(numberTable).getRow(0).getCell(0).getWidth()));
            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1)
                    .setWidthType(workbook.getTables().get(numberTable).getRow(0).getCell(0).getWidthType());

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0).getParagraphs().get(0)
                    .setIndentationLeft(0);
            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1).getParagraphs().get(0)
                    .setIndentationLeft(200);

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0).getParagraphs().get(0)
                    .setIndentationRight(0);
            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1).getParagraphs().get(0)
                    .setIndentationRight(0);

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0).getParagraphs().get(0)
                    .setIndentationFirstLine(0);

            workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1).getParagraphs().get(0)
                    .setIndentationFirstLine(0);

            XWPFRun run = workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(0).getParagraphs().get(0).createRun();
            run.removeCarriageReturn();
            run.setFontSize(9);
            run.setFontFamily("Century Gothic");
            run.setItalic(true);
            run.setText(result.get(i).get(0));
            run = workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows()-1).getCell(1).getParagraphs().get(0).createRun();
            run.setFontSize(9);
            run.setFontFamily("Century Gothic");
            if(Double.parseDouble(result.get(i).get(1)) == 0){
                run.setText("0");
            } else if (Double.parseDouble(result.get(i).get(1)) < 0.01){
                run.setText(String.format("%(.3f", Double.parseDouble(result.get(i).get(1))));
            } else if (Double.parseDouble(result.get(i).get(1)) < 0.001){
                run.setText(String.format("%(.4f", Double.parseDouble(result.get(i).get(1))));
            } else {
                run.setText(String.format("%(.2f", Double.parseDouble(result.get(i).get(1))));
            }
        }
    }

    public void saveFile(InfoList infoList, File docPath) throws IOException {
        workbook.write(new FileOutputStream(new File(docPath.getPath() + "\\" + infoList.fileName.replace(".xlsx", "") + ".docx")));
    }
}

