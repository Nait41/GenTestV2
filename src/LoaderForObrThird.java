import data.ExceptionList;
import data.InfoList;
import javafx.application.Platform;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.ls.LSOutput;

import javax.swing.*;
import java.io.*;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;

public class LoaderForObrThird extends JFrame {
    XWPFDocument workbook;
    String nameObr;
    double firstPotential, secondPotential, thirdPotential;

    public LoaderForObrThird(String nameObr) throws IOException, InvalidFormatException {
        File file = new File(Application.rootDirPath + "\\" + nameObr + ".docx");
        workbook = new XWPFDocument(new FileInputStream(file));
        this.nameObr = nameObr;
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void setFileNameForFirstFormatTable(InfoList infoList){
        XWPFRun run = workbook.getTables().get(0).getRow(11).getCell(0).addParagraph().createRun();
        run.setFontSize(11);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.fileName.replace(".xlsx", ""));
    }

    public void setBioIndexForFirstTableFormat(InfoList infoList){
        XWPFRun run = workbook.getTables().get(1).getRow(2).getCell(1).getParagraphs().get(0).createRun();
        run.setFontSize(12);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.bioIndex.get(0));
        run.setColor("ffffff");
    }

    public void setPielouEveness(InfoList infoList){
        XWPFRun run = workbook.getTables().get(1).getRow(3).getCell(1).getParagraphs().get(0).createRun();
        run.setFontSize(12);
        run.setFontFamily("Century Gothic");
        run.setText(infoList.pielouEveness);
        run.setColor("ffffff");
    }

    public void setGenusCount(InfoList infoList){
        int genusCount = 0;
        for(int i = 0; i < infoList.genus.size();i++){
            if (Double.parseDouble(infoList.genus.get(i).get(1)) != 0.0){
                genusCount++;
            }
        }
        XWPFRun run = workbook.getTables().get(1).getRow(4).getCell(1).getParagraphs().get(0).createRun();
        run.setFontSize(12);
        run.setFontFamily("Century Gothic");
        run.setText(Integer.toString(genusCount));
        run.setColor("ffffff");
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

    public void setFiveForFirstFormatTable(InfoList infoList, int numberTable){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        boolean checkBacterValue;
        for(int i = 0; i < workbook.getTables().get(numberTable).getRows().size();i++){
            checkBacterValue = true;
            if (workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 5) {
                for (int j = 0; j < genusSpecies.size(); j++) {
                    if ((genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText())
                            || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText())
                            && genusSpecies.get(j).get(0).contains("/")
                            && !(genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText() + "_")))
                            || workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals(genusSpecies.get(j).get(0)
                            .replace("_", " "))
                            || workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals(genusSpecies.get(j).get(0)
                            .replace(" ", "_"))) && checkBacterValue

                    ) {
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        if(Double.parseDouble(genusSpecies.get(j).get(1)) == 0){
                            run.setText(genusSpecies.get(j).get(1));
                        } else {
                            setValueGenusSpecies(run, genusSpecies, j);
                        }
                        checkBacterValue = false;
                    }
                }
                if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 5)
                {
                    for (int f = 0; f < workbook.getTables().get(numberTable).getRow(0).getTableCells().size();f++){
                        if(workbook.getTables().get(numberTable).getRow(0).getCell(f).getText().contains("Среднее")){
                            f++;
                            for(int k = 0; k<infoList.algs.size();k++) {
                                if (infoList.algs.get(k).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText())
                                        || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText())
                                        && infoList.algs.get(k).get(0).contains("/")
                                        && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText() + "_"))
                                        || infoList.algs.get(k).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().
                                        replace(" ", "_"))
                                        || workbook.getTables().get(numberTable).getRow(i).getCell(1).getText().equals(infoList.algs.get(k).get(0).
                                        replace("_", " "))) {
                                    if (Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.0009) {
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.6f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                                + "-" + String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.0009
                                            && Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.009) {
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                                + "-" + String.format("%(.3f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.0009
                                            && Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.09) {
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.009
                                            && Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.09) {
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.3f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.0009) {
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.009) {
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.3f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    } else {
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                        run.setFontSize(9);
                                        run.setFontFamily("Century Gothic");
                                        run.setText(String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                                + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                        workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    }
                    if(workbook.getTables().get(numberTable).getRow(i).getCell(4).getText().equals("")){
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("0.0");
                    }
                }
            }
        }
    }

    public void setSixForFirstFormatTable(InfoList infoList, int numberTable) throws XmlException, IOException {
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        for(int i = 2; i < workbook.getTables().get(numberTable).getRows().size();i++){
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 6)
                {
                    if(genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                            || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                            && genusSpecies.get(j).get(0).contains("/")
                            && !genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                            || genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().
                            replace(" ", "_"))
                            || genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().
                            replace("_", " "))
                    )
                    {
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        if(Double.parseDouble(genusSpecies.get(j).get(1)) == 0){
                            run.setText(genusSpecies.get(j).get(1));
                        } else {
                            setValueGenusSpecies(run, genusSpecies, j);
                        }
                        for(int k = 0; k<infoList.algs.size();k++)
                        {
                            if(infoList.algs.get(k).get(0).equals(genusSpecies.get(j).get(0))
                                    || (infoList.algs.get(k).get(0).contains(genusSpecies.get(j).get(0))
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(genusSpecies.get(j).get(0) + "_"))
                                    || infoList.algs.get(k).get(0).equals(genusSpecies.get(j).get(0).
                                    replace(" ", "_"))
                                    || genusSpecies.get(j).get(0).equals(infoList.algs.get(k).get(0).
                                    replace("_", " "))){
                                if (checkValueRange(infoList.algs.get(k).get(1), infoList.algs.get(k).get(2), genusSpecies.get(j).get(1)).equals("medium")){
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(5).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Среднее значение");
                                    break;
                                } else if (checkValueRange(infoList.algs.get(k).get(1), infoList.algs.get(k).get(2), genusSpecies.get(j).get(1)).equals("low")){
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(5).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Низкое значение");
                                    break;
                                } else if (checkValueRange(infoList.algs.get(k).get(1), infoList.algs.get(k).get(2), genusSpecies.get(j).get(1)).equals("high")){
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(5).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Высокое значение");
                                    break;
                                } else if (checkValueRange(infoList.algs.get(k).get(1), infoList.algs.get(k).get(2), genusSpecies.get(j).get(1)).equals("null")){
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(5).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Значение отсутствует");
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
            }
            if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 6)
            {
                for (int f = 0; f < workbook.getTables().get(numberTable).getRow(0).getTableCells().size()
                        || f < workbook.getTables().get(numberTable).getRow(1).getTableCells().size();f++) {
                    if ((workbook.getTables().get(numberTable).getRow(0).getTableCells().size() > f
                            && workbook.getTables().get(numberTable).getRow(0).getCell(f).getText().contains("Среднее"))
                            || (workbook.getTables().get(numberTable).getRow(1).getTableCells().size() > f
                            && workbook.getTables().get(numberTable).getRow(1).getCell(f).getText().contains("Среднее"))) {
                        f++;
                        for (int k = 0; k < infoList.algs.size(); k++) {
                            if (infoList.algs.get(k).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                    || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                    || infoList.algs.get(k).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().
                                    replace(" ", "_"))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0).
                                    replace("_", " "))) {
                                if (Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.0009) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.6f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.0009
                                        && Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.009) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.3f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.0009
                                        && Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.09) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.009
                                        && Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.09) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.3f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.0009) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.009) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.3f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                if(workbook.getTables().get(numberTable).getRow(i).getCell(4).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(4).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("0.0");
                }
                if(workbook.getTables().get(numberTable).getRow(i).getCell(5).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(5).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("Значение отсутствует");
                }
            }
        }
    }

    public void setFourFormat(InfoList infoList, int numberTable){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        double potentialLow = 0, potentialMedium = 0, potentialHigh = 0, potentialNull = 0;
        for(int i = 1; i < workbook.getTables().get(numberTable).getRows().size();i++){
            if (workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 2){
                if (workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().contains("ПОТЕНЦИАЛ")){
                    /*
                    double resSum = (1*potentialNull+2*potentialLow+3*potentialMedium+4*potentialHigh)/(potentialNull+potentialLow+potentialMedium+potentialHigh);
                    if (resSum < 1.75){
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("ОЧЕНЬ НИЗКИЙ");
                    } else if (resSum < 2.5){
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("НИЗКИЙ");
                    } else if (resSum < 3.25){
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("СРЕДНИЙ");
                    } else {
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        run.setText("ВЫСОКИЙ");
                    }
                    potentialNull = 0;
                    potentialLow = 0;
                    potentialMedium = 0;
                    potentialHigh = 0;
                    */
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(1).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("СРЕДНИЙ");
                }
            }
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 4)
                {
                    if(genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                            || (genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                            && genusSpecies.get(j).get(0).contains("/")
                            && !genusSpecies.get(j).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                            || genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().
                            replace(" ", "_"))
                            || genusSpecies.get(j).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().
                            replace("_", " "))
                    )
                    {
                        XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Century Gothic");
                        if(Double.parseDouble(genusSpecies.get(j).get(1)) == 0){
                            run.setText(genusSpecies.get(j).get(1));
                        } else {
                            setValueGenusSpecies(run, genusSpecies, j);
                        }
                        for(int k = 0; k<infoList.algs.size();k++)
                        {
                            if(infoList.algs.get(k).get(0).equals(genusSpecies.get(j).get(0))
                                    || (infoList.algs.get(k).get(0).contains(genusSpecies.get(j).get(0))
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(genusSpecies.get(j).get(0) + "_"))
                                    || infoList.algs.get(k).get(0).equals(genusSpecies.get(j).get(0).
                                    replace(" ", "_"))
                                    || genusSpecies.get(j).get(0).equals(infoList.algs.get(k).get(0).
                                    replace("_", " "))){
                                if (checkValueRange(infoList.algs.get(k).get(1), infoList.algs.get(k).get(2), genusSpecies.get(j).get(1)).equals("medium")){
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Среднее значение");
                                    potentialMedium++;
                                    break;
                                } else if (checkValueRange(infoList.algs.get(k).get(1), infoList.algs.get(k).get(2), genusSpecies.get(j).get(1)).equals("low")){
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Низкое значение");
                                    potentialLow++;
                                    break;
                                } else if (checkValueRange(infoList.algs.get(k).get(1), infoList.algs.get(k).get(2), genusSpecies.get(j).get(1)).equals("high")){
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Высокое значение");
                                    potentialHigh++;
                                    break;
                                } else if (checkValueRange(infoList.algs.get(k).get(1), infoList.algs.get(k).get(2), genusSpecies.get(j).get(1)).equals("null")){
                                    run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText("Значение отсутствует");
                                    potentialNull++;
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
            }
            if(workbook.getTables().get(numberTable).getRow(i).getTableCells().size() == 4)
            {
                for (int f = 0; f < workbook.getTables().get(numberTable).getRow(0).getTableCells().size();f++) {
                    if (workbook.getTables().get(numberTable).getRow(0).getCell(f).getText().contains("Среднее")) {
                        if (workbook.getTables().get(numberTable).getRow(0).getTableCells().size() >
                                workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows() - 1)
                                        .getTableCells().size() && workbook.getTables().get(numberTable).getRow(workbook.getTables().get(numberTable).getNumberOfRows() - 1)
                                .getTableCells().size() != 2){
                            f--;
                        }
                        for (int k = 0; k < infoList.algs.size(); k++) {
                            if (infoList.algs.get(k).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                    || (infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText())
                                    && infoList.algs.get(k).get(0).contains("/")
                                    && !infoList.algs.get(k).get(0).contains(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText() + "_"))
                                    || infoList.algs.get(k).get(0).equals(workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().
                                    replace(" ", "_"))
                                    || workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals(infoList.algs.get(k).get(0).
                                    replace("_", " "))) {
                                if (Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.0009) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.6f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.0009
                                        && Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.009) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.3f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.0009
                                        && Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.09) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.009
                                        && Double.parseDouble(infoList.algs.get(k).get(2)) <= 0.09) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.3f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.0009) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.4f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else if (Double.parseDouble(infoList.algs.get(k).get(1)) <= 0.009) {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.3f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                } else {
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).removeParagraph(0);
                                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(f).addParagraph().createRun();
                                    run.setFontSize(9);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(1)))
                                            + "-" + String.format("%(.2f", Double.parseDouble(infoList.algs.get(k).get(2))));
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                                    workbook.getTables().get(numberTable).getRow(i).getCell(f).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                    break;
                                }
                            }
                        }
                        break;
                    }
                }
                if (!workbook.getTables().get(numberTable).getRow(i).getCell(2).getText().equals("")
                        &&  workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals("")
                        &&  !workbook.getTables().get(numberTable).getRow(i).getCell(0).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("Интерпретация не определена");
                    run.setColor("a60000");
                }
                if(workbook.getTables().get(numberTable).getRow(i).getCell(2).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(2).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("0.0");
                    potentialLow++;
                }
                if(workbook.getTables().get(numberTable).getRow(i).getCell(3).getText().equals("")){
                    XWPFRun run = workbook.getTables().get(numberTable).getRow(i).getCell(3).getParagraphs().get(0).createRun();
                    run.setFontSize(9);
                    run.setFontFamily("Century Gothic");
                    run.setText("Значение отсутствует");
                }
            }
        }
    }

    public void setAddition(InfoList infoList, int numberAddition){
        ArrayList<ArrayList<String>> genusSpecies = new ArrayList<>();
        genusSpecies.addAll(infoList.genus);
        genusSpecies.addAll(infoList.species);
        int additionNumber = 0;
        int parNumber = 0;
        int counter = 0;
        for(int i = 0; i < workbook.getParagraphs().size();i++)
        {
            if(workbook.getParagraphs().get(i).getText().equals("Дополнение") || workbook.getParagraphs().get(i).getText().contains("ДОПОЛНЕНИЕ")){
                parNumber = i;
                additionNumber++;
                if(additionNumber == numberAddition){
                    break;
                }
            }
        }

        CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
        cTAbstractNum.setAbstractNumId(BigInteger.valueOf(50 + numberAddition - 1));

        CTLvl cTLvl = cTAbstractNum.addNewLvl();
        cTLvl.setIlvl(BigInteger.valueOf(0));
        cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
        cTLvl.addNewLvlText().setVal("%1.");
        cTLvl.addNewStart().setVal(BigInteger.valueOf(1));
        cTLvl.addNewRPr();
        cTLvl.getRPr().addNewSz().setVal(9*2);
        cTLvl.getRPr().addNewSzCs().setVal(9*2);
        cTLvl.getRPr().addNewRFonts().setAscii("Century Gothic");

        XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
        XWPFNumbering numbering = workbook.createNumbering();
        BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
        BigInteger numID = numbering.addNum(abstractNumID);

        parNumber+=1;
        boolean checkFirstBacteria = true;
        for(int d = 0; d < infoList.uniqBact.size(); d++)
        {
            for (int j = 0; j < genusSpecies.size(); j++)
            {
                for(int k = 0; k < infoList.algs.size();k++)
                {
                    if(infoList.uniqBact.get(d).equals(infoList.algs.get(k).get(7)))
                    {
                        String firstMediumValue;
                        String secondMediumValue;
                        if(infoList.algs.get(k).get(3).equals("")){
                            firstMediumValue = infoList.algs.get(k).get(1);
                            secondMediumValue = infoList.algs.get(k).get(2);
                        } else if(Double.parseDouble(infoList.algs.get(k).get(3)) == 0){
                            firstMediumValue = "0";
                            secondMediumValue = "0";
                        } else {
                            firstMediumValue = infoList.algs.get(k).get(1);
                            secondMediumValue = infoList.algs.get(k).get(3);
                        }
                        if(infoList.algs.get(k).get(0).equals(genusSpecies.get(j).get(0))){
                            if (checkValueRange(firstMediumValue, secondMediumValue, genusSpecies.get(j).get(1)).equals("low")
                                    && !infoList.algs.get(k).get(5).equals("0.0")){
                                if(checkFirstBacteria)
                                {
                                    XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                    XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                    xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                    xwpfParagraph.setIndentationLeft(420);
                                    XWPFRun run = xwpfParagraph.createRun();
                                    run.setFontSize(9);
                                    run.setBold(true);
                                    run.setItalic(true);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.uniqBact.get(d));
                                    run.addBreak();
                                    counter++;
                                    checkFirstBacteria = false;
                                }
                                XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                xwpfParagraph.setAlignment(ParagraphAlignment.BOTH);
                                xwpfParagraph.setIndentationLeft(0);
                                xwpfParagraph.setFirstLineIndent(420);
                                xwpfParagraph.setNumID(numID);
                                XWPFRun run = xwpfParagraph.createRun();
                                run.setFontSize(9);
                                run.setItalic(true);
                                run.setFontFamily("Century Gothic");
                                run.setText(infoList.algs.get(k).get(5));
                                run.addBreak();
                                counter++;
                                break;
                            } else if (checkValueRange(firstMediumValue, secondMediumValue, genusSpecies.get(j).get(1)).equals("null")
                                    && !infoList.algs.get(k).get(4).equals("0.0")){
                                if(checkFirstBacteria)
                                {
                                    XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                    XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                    xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                    xwpfParagraph.setIndentationLeft(420);
                                    XWPFRun run = xwpfParagraph.createRun();
                                    run.setFontSize(9);
                                    run.setBold(true);
                                    run.setItalic(true);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.uniqBact.get(d));
                                    run.addBreak();
                                    counter++;
                                    checkFirstBacteria = false;
                                }
                                XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                xwpfParagraph.setAlignment(ParagraphAlignment.BOTH);
                                xwpfParagraph.setIndentationLeft(0);
                                xwpfParagraph.setFirstLineIndent(420);
                                xwpfParagraph.setNumID(numID);
                                XWPFRun run = xwpfParagraph.createRun();
                                run.setFontSize(9);
                                run.setItalic(true);
                                run.setFontFamily("Century Gothic");
                                run.setText(infoList.algs.get(k).get(4));
                                run.addBreak();
                                counter++;
                                break;
                            } else if (checkValueRange(firstMediumValue, secondMediumValue, genusSpecies.get(j).get(1)).equals("high")
                                    && !infoList.algs.get(k).get(6).equals("0.0")){
                                if(checkFirstBacteria)
                                {
                                    XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                    XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                    xwpfParagraph.setAlignment(ParagraphAlignment.LEFT);
                                    xwpfParagraph.setIndentationLeft(420);
                                    XWPFRun run = xwpfParagraph.createRun();
                                    run.setFontSize(9);
                                    run.setBold(true);
                                    run.setItalic(true);
                                    run.setFontFamily("Century Gothic");
                                    run.setText(infoList.uniqBact.get(d));
                                    run.addBreak();
                                    counter++;
                                    checkFirstBacteria = false;
                                }
                                XmlCursor xmlCursor = workbook.getParagraphs().get(parNumber+counter).getCTP().newCursor();
                                XWPFParagraph xwpfParagraph = workbook.insertNewParagraph(xmlCursor);
                                xwpfParagraph.setAlignment(ParagraphAlignment.BOTH);
                                xwpfParagraph.setIndentationLeft(0);
                                xwpfParagraph.setFirstLineIndent(420);
                                xwpfParagraph.setNumID(numID);
                                XWPFRun run = xwpfParagraph.createRun();
                                run.setFontSize(9);
                                run.setItalic(true);
                                run.setFontFamily("Century Gothic");
                                run.setText(infoList.algs.get(k).get(6));
                                run.addBreak();
                                counter++;
                                break;
                            }
                        }
                    }
                }
            }
            checkFirstBacteria = true;
        }
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

    String checkValueRange(String firstMediumValue, String secondMediumValue, String currentValue){
        if(Double.parseDouble(currentValue) >= Double.parseDouble(firstMediumValue) && Double.parseDouble(currentValue) <= Double.parseDouble(secondMediumValue)){
            return "medium";
        } else if (Double.parseDouble(currentValue) < Double.parseDouble(firstMediumValue) && Double.parseDouble(currentValue) > 0){
            return "low";
        } else if (Double.parseDouble(currentValue) > Double.parseDouble(secondMediumValue)){
            return "high";
        } else if (Double.parseDouble(currentValue) == 0){
            return "null";
        }
        return "null";
    }

    public void saveFile(InfoList infoList, File docPath) throws IOException {
        workbook.write(new FileOutputStream(new File(docPath.getPath() + "\\" + infoList.fileName.replace(".xlsx", "") + ".docx")));
    }
}

