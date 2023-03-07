import data.ExceptionList;
import data.InfoList;
import fileView.XLXSOpen;
import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.Tooltip;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.AnchorPane;
import javafx.scene.shape.Circle;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import javafx.util.Duration;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.ResourceBundle;

public class MainController {
    public static boolean hintsOption = true;
    public static boolean descriptionOption = false;
    public static boolean genusOption = false;
    public static boolean mediumRangeOption = false;
    public static boolean missingOption = false;
    public static boolean exceptCheck = false;
    public InfoList infoList;
    AlgOpen alg;
    AlgsUrogenital algsUrogenital;
    ArrayList<String> content_list = new ArrayList<>();
    List<File> samplePath;
    String selectedSample = "";
    String selectedException = "";
    MainLoader docLoad;
    LoaderForObrSecond loaderForObrSecond;
    LoaderForObrFirst loaderForObrFirst;
    LoaderForObrThird loaderForObrThird;
    LoaderForObrFour loaderForObrFour;
    LoaderForObrFive loaderForObrFive;
    LoaderForObrSix loaderForObrSix;
    LoaderForObrSeven loaderForObrSeven;
    XLXSOpen xlxsOpen;
    File saveSampleDir;
    boolean checkLoad, checkUnload, checkStart = false;
    int counter, counter_files;
    public static String errorMessageStr = "";

    @FXML
    private ResourceBundle resources;

    @FXML
    private URL location;

    @FXML
    private Button dirLoadButton;

    @FXML
    private Button algsTable;

    @FXML
    private Button dirUnloadButton;

    ArrayList<String> langs = new ArrayList<>();

    @FXML
    private ListView<String> listSample;

    @FXML
    private Text loadStatus;

    @FXML
    private Text loadStatus_end;

    @FXML
    private Text loadStatusFileNumber;

    @FXML
    private Button startButton;

    @FXML
    public Label lowLoadText = new Label("");

    @FXML
    private AnchorPane mainPanel;

    @FXML
    public Button closeButton;

    @FXML
    private Button exceptionButton;

    @FXML
    private AnchorPane exceptPane;

    @FXML
    private ListView<String> exceptView;

    @FXML
    private Button sampleEditButton;

    @FXML
    private Button options;

    public MainController() throws IOException, InvalidFormatException {
    }

    void feelLangs(){
        langs.add("Рассширенный образец");
        langs.add("Шаблон расширенный (дети)");
        langs.add("Краткая версия урогенитального микробиома");
        langs.add("Шаблон стандарт");
        langs.add("Шаблон патогенная нагрузка");
    }

    int getCounter(int rowCount, int currentNumber) {
        Double temp = new Double(100/rowCount);
        return temp.intValue() + currentNumber;
    }

    void feelExceptLangs(){
        if (!exceptView.getItems().contains("Не для всех бактерий определены среднии значения популяции"))
        {
            if(exceptCheck && mediumRangeOption){
                exceptView.getItems().add("Не для всех бактерий определены среднии значения популяции");
            }
        }
        if(!exceptView.getItems().contains("Не для всех бактерий определен род")){
            if(GenusExceptionAnalyzer.genusException && genusOption){
                exceptView.getItems().add("Не для всех бактерий определен род");
            }
        }
        if(!exceptView.getItems().contains("Не все бактерии описаны")){
            if(DescriptionExceptionAnalyzer.descriptionExcept && descriptionOption){
                exceptView.getItems().add("Не все бактерии описаны");
            }
        }
        if(!exceptView.getItems().contains("Список отсутствующих бактерий в образце")){
            if(DescriptionExceptionAnalyzer.descriptionExcept && missingOption){
                exceptView.getItems().add("Список отсутствующих бактерий в образце");
            }
        }
    }

    public void addHinds(){
        Tooltip tipSampleEdit = new Tooltip();
        tipSampleEdit.setText("Нажмите, для того, чтобы перейти к меню изменения шаблонов");
        tipSampleEdit.setStyle("-fx-text-fill: turquoise;");
        sampleEditButton.setTooltip(tipSampleEdit);

        Tooltip tipAlgsTable = new Tooltip();
        tipAlgsTable.setText("Нажмите, для того, чтобы перейти к редактированию таблицы алгоритмов");
        tipAlgsTable.setStyle("-fx-text-fill: turquoise;");
        algsTable.setTooltip(tipAlgsTable);

        Tooltip tipLoad = new Tooltip();
        tipLoad.setText("Выберите папку, в которой находятся xlsx файлы");
        tipLoad.setStyle("-fx-text-fill: turquoise;");
        dirLoadButton.setTooltip(tipLoad);

        Tooltip tipOptions = new Tooltip();
        tipOptions.setText("Нажмите, для того, чтобы перейти в опции");
        tipOptions.setStyle("-fx-text-fill: turquoise;");
        options.setTooltip(tipOptions);

        Tooltip tipUnLoad = new Tooltip();
        tipUnLoad.setText("Выберите папку, в которую должны сохраняться готовые отчеты");
        tipUnLoad.setStyle("-fx-text-fill: turquoise;");
        dirUnloadButton.setTooltip(tipUnLoad);

        Tooltip tipStart = new Tooltip();
        tipStart.setText("Нажмите, для того, чтобы получить готовые отчеты");
        tipStart.setStyle("-fx-text-fill: turquoise;");
        startButton.setTooltip(tipStart);

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы закрыть приложение");
        closeStart.setStyle("-fx-text-fill: turquoise;");
        closeButton.setTooltip(closeStart);

        Tooltip exceptionTip = new Tooltip();
        exceptionTip.setText("Нажмите на кнопку, чтобы посмотреть список проблем");
        exceptionTip.setStyle("-fx-text-fill: turquoise;");
        exceptionButton.setTooltip(exceptionTip);
    }

    public void removeHinds(){
        algsTable.setTooltip(null);
        dirLoadButton.setTooltip(null);
        options.setTooltip(null);
        dirUnloadButton.setTooltip(null);
        startButton.setTooltip(null);
        closeButton.setTooltip(null);
        exceptionButton.setTooltip(null);
    }

    public static boolean tempHints = true;

    @FXML
    void initialize() throws FileNotFoundException, InterruptedException {
        Timeline timeline = new Timeline(new KeyFrame(Duration.seconds(3), e -> {
            if (tempHints != hintsOption){
                tempHints = hintsOption;
                if (hintsOption == true){
                    addHinds();
                } else
                {
                    removeHinds();
                }
            }
            if (!mediumRangeOption){
                if(ExceptionList.exceptBact == null){
                    System.out.println(1);
                    if(exceptView.getItems().contains("Не для всех бактерий определены среднии значения популяции")) {
                        exceptView.getItems().remove("Не для всех бактерий определены среднии значения популяции");
                    }
                }
            }
            if (!genusOption){
                if(ExceptionList.genusExceptBact == null){
                    if(exceptView.getItems().contains("Не для всех бактерий определен род")) {
                        exceptView.getItems().remove("Не для всех бактерий определен род");
                    }
                }
            }
            if (!descriptionOption){
                if(ExceptionList.descriptionExpect == null){
                    if(exceptView.getItems().contains("Не все бактерии описаны")) {
                        exceptView.getItems().remove("Не все бактерии описаны");
                    }
                }
            }
            if (!mediumRangeOption){
                if(exceptView.getItems().contains("Не для всех бактерий определены среднии значения популяции")) {
                    exceptView.getItems().remove("Не для всех бактерий определены среднии значения популяции");
                }
            }
            if (!genusOption){
                if(exceptView.getItems().contains("Не для всех бактерий определен род")) {
                    exceptView.getItems().remove("Не для всех бактерий определен род");
                }
            }
            if (!descriptionOption){
                if(exceptView.getItems().contains("Не все бактерии описаны")) {
                    exceptView.getItems().remove("Не все бактерии описаны");
                }
            }
            if (!missingOption){
                if(exceptView.getItems().contains("Список отсутствующих бактерий в образце")) {
                    exceptView.getItems().remove("Список отсутствующих бактерий в образце");
                }
            }
            if (!mediumRangeOption && !descriptionOption && !genusOption && !missingOption)
            {
                exceptionButton.setVisible(false);
                exceptPane.setVisible(false);
            }
        }));
        timeline.setCycleCount(-1);
        timeline.play();
        addHinds();
        exceptPane.setVisible(false);
        exceptionButton.setVisible(false);

        FileInputStream sampleEditStream = new FileInputStream(Application.rootDirPath +"\\sampleEdit.png");
        Image sampleEditImage = new Image(sampleEditStream);
        ImageView sampleEditView = new ImageView(sampleEditImage);
        sampleEditButton.graphicProperty().setValue(sampleEditView);

        FileInputStream optionsStream = new FileInputStream(Application.rootDirPath + "\\options.png");
        Image optionsImage = new Image(optionsStream);
        ImageView optionsView = new ImageView(optionsImage);
        options.graphicProperty().setValue(optionsView);

        FileInputStream loadStream = new FileInputStream(Application.rootDirPath + "\\load.png");
        Image loadImage = new Image(loadStream);
        ImageView loadView = new ImageView(loadImage);
        dirLoadButton.graphicProperty().setValue(loadView);

        FileInputStream unloadStream = new FileInputStream(Application.rootDirPath + "\\unload.png");
        Image unloadImage = new Image(unloadStream);
        ImageView unloadView = new ImageView(unloadImage);
        dirUnloadButton.graphicProperty().setValue(unloadView);

        FileInputStream startStream = new FileInputStream(Application.rootDirPath + "\\start.png");
        Image startImage = new Image(startStream);
        ImageView startView = new ImageView(startImage);
        startButton.graphicProperty().setValue(startView);

        FileInputStream closeStream = new FileInputStream(Application.rootDirPath + "\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);

        FileInputStream exceptionStream = new FileInputStream(Application.rootDirPath + "\\exception.png");
        Image exceptionImage = new Image(exceptionStream);
        ImageView exceptionv = new ImageView(exceptionImage);
        exceptionButton.graphicProperty().setValue(exceptionv);

        FileInputStream algsTableStream = new FileInputStream(Application.rootDirPath + "\\algsTable.png");
        Image algsTableImage = new Image(algsTableStream);
        ImageView algsTableView = new ImageView(algsTableImage);
        algsTable.graphicProperty().setValue(algsTableView);

        algsTable.setOnAction(ActionEvent -> {
            AlgsTableController algsTableController = new AlgsTableController();
            try {
                algsTableController.start(new Stage());
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        sampleEditButton.setOnAction(ActionEvent -> {
            ErrorController errorController = new ErrorController();
            try {
                errorMessageStr = "Данная опция пока что отсутствует";
                errorController.start(new Stage());
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        exceptView.getSelectionModel().selectedItemProperty().addListener(new ChangeListener<String>() {
            ExceptionAnalyzer exceptionAnalyzer = new ExceptionAnalyzer();
            GenusExceptionAnalyzer genusExceptionAnalyzer = new GenusExceptionAnalyzer();
            DescriptionExceptionAnalyzer descriptionExceptionAnalyzer = new DescriptionExceptionAnalyzer();
            @Override
            public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
                selectedException = exceptView.getSelectionModel().getSelectedItem();
                if(selectedException.equals("Не для всех бактерий определены среднии значения популяции")){
                    try {
                        exceptionAnalyzer.start(new Stage());
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if(selectedException.equals("Не для всех бактерий определен род")){
                    try {
                        genusExceptionAnalyzer.start(new Stage());
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if(selectedException.equals("Не все бактерии описаны")) {
                    try {
                        descriptionExceptionAnalyzer.start(new Stage());
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if(selectedException.equals("Список отсутствующих бактерий в образце")){
                }
            }
        });

        int r = 60;
        startButton.setShape(new Circle(r));
        startButton.setMinSize(r*2, r*2);
        startButton.setMaxSize(r*2, r*2);

        checkLoad = false;
        checkUnload = false;
        feelLangs();
        listSample.getItems().addAll(langs);
        listSample.getSelectionModel().selectedItemProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observableValue, String s, String t1) {
                selectedSample = listSample.getSelectionModel().getSelectedItem();
            }
        });

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        exceptionButton.setOnAction(actionEvent -> {
            if(exceptPane.isVisible()){
                exceptPane.setVisible(false);
            }
            else{
                exceptPane.setVisible(true);
                feelExceptLangs();
                if (exceptView.getItems().size()<2)
                {
                    ExceptionAnalyzer exceptionAnalyzer = new ExceptionAnalyzer();
                    GenusExceptionAnalyzer genusExceptionAnalyzer = new GenusExceptionAnalyzer();
                    DescriptionExceptionAnalyzer descriptionExceptionAnalyzer = new DescriptionExceptionAnalyzer();
                    if(selectedException.equals("Не для всех бактерий определены среднии значения популяции")){
                        try {
                            exceptionAnalyzer.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    if(selectedException.equals("Не для всех бактерий определен род")){
                        try {
                            genusExceptionAnalyzer.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    if(selectedException.equals("Не все бактерии описаны")){
                        try {
                            descriptionExceptionAnalyzer.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    if(selectedException.equals("Список отсутствующих бактерий в образце")){
                    }
                }
            }
        });

        dirLoadButton.setOnAction(actionEvent -> {
            if(!checkStart)
            {
                loadStatus.setText("");
                loadStatus_end.setText("");
                loadStatusFileNumber.setText("");
                DirectoryChooser directoryChooser = new DirectoryChooser();
                File dir = directoryChooser.showDialog(new Stage());
                File[] file = dir.listFiles();
                samplePath = Arrays.asList(file);
                checkLoad = true;
            }
            else
            {
                errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });

        options.setOnAction(ActionEvent -> {
            OptionsController optionsController = new OptionsController();
            try {
                optionsController.start(new Stage());
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        dirUnloadButton.setOnAction(actionEvent -> {
                    if(!checkStart)
                    {
                        loadStatus.setText("");
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        DirectoryChooser directoryChooser = new DirectoryChooser();
                        saveSampleDir = directoryChooser.showDialog(new Stage());
                        checkUnload = true;

                    }
                    else
                    {
                        errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
        startButton.setOnAction(actionEvent -> {
                    if(!checkStart){
                        loadStatus.setText("");
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        if(checkLoad & checkUnload){
                            if(!selectedSample.equals(""))
                            {
                                if(samplePath.size() != 0)
                                {
                                    if (!MainController.mediumRangeOption && !MainController.descriptionOption
                                            && !MainController.genusOption && !MainController.missingOption){
                                        exceptionButton.setVisible(false);
                                    }
                                    checkStart = true;
                                    ExceptionList.exceptBact = new ArrayList<>();
                                    ExceptionList.genusExceptBact = new ArrayList<>();
                                    ExceptionList.descriptionExpect = new ArrayList<>();
                                    if(selectedSample.equals("Рассширенный образец")){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                ArrayList<String> potentials = new ArrayList<>();
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            loaderForObrSecond = new LoaderForObrSecond("obr");
                                                            loaderForObrFour = new LoaderForObrFour("obr_3");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getPhylum(infoList);
                                                            xlxsOpen.getGenus(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            xlxsOpen.getSpecies(infoList);
                                                            xlxsOpen.getBioIndex(infoList);
                                                            xlxsOpen.getPielouEveness(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        loaderForObrSecond.setFileNameForFirstFormatTable(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        if (infoList.bioIndex.size() != 0){
                                                            loaderForObrSecond.setBioIndexForFirstTableFormat(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        }
                                                        if (infoList.pielouEveness != null){
                                                            loaderForObrSecond.setPielouEveness(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        }
                                                        loaderForObrSecond.setGenusCount(infoList);
                                                        if (infoList.phylum.size() != 0){
                                                            loaderForObrSecond.setRatioPhylum(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        }
                                                        loaderForObrSecond.setRatioSpecies(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        try {
                                                            loaderForObrSecond.setThreeDoubleFormat(infoList, 5);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        loaderForObrSecond.setFiveForFirstFormatTable(infoList, 6);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        loaderForObrSecond.setFiveForFirstFormatTable(infoList, 7);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        loaderForObrSecond.setFiveForFirstFormatTable(infoList, 8);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        loaderForObrSecond.setFourFormat(infoList, 8);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        loaderForObrSecond.setFourFormat(infoList, 9);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        loaderForObrSecond.setFourFormat(infoList, 10);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        loaderForObrSecond.setFourFormat(infoList, 11);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        try {
                                                            if (infoList.phylum.size() != 0){
                                                                loaderForObrSecond.setPhylum(infoList);
                                                                loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            }
                                                            loaderForObrSecond.setFiveForSecondFormatTable(infoList, 12, true, false);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.setFiveForSecondFormatTable(infoList, 13, true, false);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.setFiveForSecondFormatTable(infoList, 14, true, false);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.setFiveForSecondFormatTable(infoList, 15, true, false);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.setFiveForSecondFormatTable(infoList, 16, true, false);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.setFiveForSecondFormatTable(infoList, 17, true, false);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.setFiveForSecondFormatTable(infoList, 18, true, false);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.setSixForFirstFormatTable(infoList, 19);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.setTwoFormatWithSer(infoList,23, "genus");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.setTwoFormatWithSer(infoList,24, "species");;
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        } catch (XmlException e) {
                                                            e.printStackTrace();
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        } catch (ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        loaderForObrSecond.setAddition(infoList, 1);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        potentials = loaderForObrSecond.getPotentials();
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        try {
                                                            loaderForObrSecond.saveFile(infoList, saveSampleDir);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrSecond.getClose();
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        loaderForObrFour.setFileNameForSecondFormatTable(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");

                                                        if (infoList.bioIndex.size() != 0) {
                                                            loaderForObrFour.setBioindexInLowInfo(infoList);
                                                        }
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        if (infoList.pielouEveness != null) {
                                                            loaderForObrFour.setPielouEvenessInLowInfo(infoList);
                                                        }
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        loaderForObrFour.setGenusCountInLowInfo(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        if (infoList.phylum.size() != 0) {
                                                            loaderForObrFour.setPhylumRatioInLowInfo(infoList);
                                                        }
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        loaderForObrFour.setSpeciesRatioInLowInfo(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        try {
                                                            loaderForObrFour.setSixForFirstFormatTable(infoList, 2);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrFour.setFiveForFirstFormatTable(infoList, 3);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrFour.setFiveForFirstFormatTable(infoList, 4);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrFour.setFiveForFirstFormatTable(infoList, 5);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrFour.setPotentialForSecondTable(6, potentials);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrFour.setAddition(infoList, 1);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrFour.setTwoFormatWithSer(infoList, 8, "genus");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrFour.setTwoFormatWithSer(infoList, 9, "species");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                        } catch (XmlException e) {
                                                            e.printStackTrace();
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            loaderForObrFour.saveFile(infoList, saveSampleDir);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            loaderForObrFour.getClose();
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(49, counter)) + " %");
                                                            xlxsOpen.getClose();
                                                            loadStatus.setText("Загрузка: 100%");
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                                if(mediumRangeOption || genusOption || descriptionOption || missingOption){
                                                    exceptionButton.setVisible(true);
                                                }
                                            }
                                        }.start();
                                    } else if(selectedSample.equals("Краткая версия урогенитального микробиома")){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++) {
                                                    if(samplePath.get(i).getPath().contains(".xlsx"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            loaderForObrFirst = new LoaderForObrFirst("obr_1");
                                                            algsUrogenital = new AlgsUrogenital(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getFamily(infoList);
                                                            xlxsOpen.getPhylum(infoList);
                                                            xlxsOpen.getGenus(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            xlxsOpen.getSpecies(infoList);
                                                            xlxsOpen.getBioIndex(infoList);
                                                            loaderForObrFirst.setFileNameForSecond(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(8, counter)) + " %");
                                                            if (infoList.bioIndex.size() != 0){
                                                                loaderForObrFirst.setBioIndex(infoList, 0);
                                                            }
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(8, counter)) + " %");
                                                            loaderForObrFirst.setDataInFiveColumnTable(infoList, 0);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(8, counter)) + " %");
                                                            loaderForObrFirst.setTwoFormatWithSer(infoList, 1, "genus");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(8, counter)) + " %");
                                                            loaderForObrFirst.setTwoFormatWithSer(infoList, 2, "family");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(8, counter)) + " %");
                                                            loaderForObrFirst.setTwoFormatWithSer(infoList, 3, "species");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(8, counter)) + " %");
                                                            loaderForObrFirst.saveFile(infoList,saveSampleDir);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(8, counter)) + " %");
                                                            try {
                                                                loaderForObrFirst.getClose();
                                                                loadStatus.setText("Загрузка: 100 %");
                                                                xlxsOpen.getClose();
                                                            } catch (IOException e) {
                                                                e.printStackTrace();
                                                            }
                                                        } catch (IOException | XmlException | ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                                if(mediumRangeOption || genusOption || descriptionOption || missingOption){
                                                    exceptionButton.setVisible(true);
                                                }
                                            }
                                        }.start();
                                    } else if (selectedSample.equals("Шаблон расширенный (дети)")){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            loaderForObrThird = new LoaderForObrThird("obr_2");
                                                            loaderForObrFive = new LoaderForObrFive("obr_4");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getPhylum(infoList);
                                                            xlxsOpen.getGenus(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            xlxsOpen.getSpecies(infoList);
                                                            xlxsOpen.getBioIndex(infoList);
                                                            xlxsOpen.getPielouEveness(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        loaderForObrThird.setFileNameForFirstFormatTable(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        if (infoList.bioIndex.size() != 0){
                                                            loaderForObrThird.setBioIndexForFirstTableFormat(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        }
                                                        if (infoList.pielouEveness != null){
                                                            loaderForObrThird.setPielouEveness(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        }
                                                        loaderForObrThird.setGenusCount(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        loaderForObrThird.setFiveForFirstFormatTable(infoList, 2);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        loaderForObrThird.setFiveForFirstFormatTable(infoList, 3);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        loaderForObrThird.setFiveForFirstFormatTable(infoList, 4);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        loaderForObrThird.setFourFormat(infoList, 5);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        loaderForObrThird.setFourFormat(infoList, 6);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        try {
                                                            loaderForObrThird.setSixForFirstFormatTable(infoList, 7);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        } catch (XmlException e) {
                                                            e.printStackTrace();
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            loaderForObrThird.setTwoFormatWithSer(infoList, 11, "genus");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            loaderForObrThird.setTwoFormatWithSer(infoList, 12, "species");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        loaderForObrThird.setAddition(infoList, 1);
                                                        try {
                                                            loaderForObrThird.saveFile(infoList, saveSampleDir);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            loaderForObrThird.getClose();
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        loaderForObrFive.setFileNameForSecondFormatTable(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        if (infoList.bioIndex.size() != 0) {
                                                            loaderForObrFive.setBioindexInLowInfo(infoList);
                                                        }
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        if (infoList.pielouEveness != null) {
                                                            loaderForObrFive.setPielouEvenessInLowInfo(infoList);
                                                        }
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        loaderForObrFive.setGenusCountInLowInfo(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        try {
                                                            loaderForObrFive.setSixForFirstFormatTable(infoList, 2);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            loaderForObrFive.setFiveForFirstFormatTable(infoList, 3);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            loaderForObrFive.setFiveForFirstFormatTable(infoList, 4);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            loaderForObrFive.setFiveForFirstFormatTable(infoList, 5);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            loaderForObrFive.setAddition(infoList, 1);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            loaderForObrFive.setTwoFormatWithSer(infoList, 7, "genus");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            loaderForObrFive.setTwoFormatWithSer(infoList, 8, "species");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                        } catch (XmlException e) {
                                                            e.printStackTrace();
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            loaderForObrFive.saveFile(infoList, saveSampleDir);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            loaderForObrFive.getClose();
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(32, counter)) + " %");
                                                            xlxsOpen.getClose();
                                                            loadStatus.setText("Загрузка: 100%");
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                            }
                                        }.start();
                                    } else if (selectedSample.equals("Шаблон стандарт")){
                                    new Thread(){
                                        @Override
                                        public void run(){
                                            counter_files = 0;
                                            for (int i = 0; i<samplePath.size();i++)
                                            {
                                                if(samplePath.get(i).getPath().contains(".xlsx"))
                                                {
                                                    loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                    counter = 0;
                                                    infoList = new InfoList();
                                                    try {
                                                        xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                        loaderForObrSix = new LoaderForObrSix("obr_5");
                                                        alg = new AlgOpen(infoList);
                                                    } catch (IOException e) {
                                                        e.printStackTrace();
                                                    } catch (InvalidFormatException e) {
                                                        e.printStackTrace();
                                                    }
                                                    try {
                                                        xlxsOpen.getPhylum(infoList);
                                                        xlxsOpen.getGenus(infoList);
                                                        xlxsOpen.getFileName(infoList);
                                                        xlxsOpen.getSpecies(infoList);
                                                        xlxsOpen.getBioIndex(infoList);
                                                        xlxsOpen.getPielouEveness(infoList);
                                                    } catch (IOException e) {
                                                        e.printStackTrace();
                                                    }
                                                    loaderForObrSix.setFileNameForFirstFormatTable(infoList);
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    if (infoList.bioIndex.size() != 0){
                                                        loaderForObrSix.setBioIndexForFirstTableFormat(infoList);
                                                    }
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    if (infoList.pielouEveness != null){
                                                        loaderForObrSix.setPielouEveness(infoList);
                                                    }
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    loaderForObrSix.setGenusCount(infoList);
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    if (infoList.phylum.size() != 0){
                                                        loaderForObrSix.setRatioPhylum(infoList);
                                                    }
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    loaderForObrSix.setRatioSpecies(infoList);
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    try {
                                                        loaderForObrSix.setThreeDoubleFormat(infoList, 5);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    } catch (IOException e) {
                                                        e.printStackTrace();
                                                    } catch (InvalidFormatException e) {
                                                        e.printStackTrace();
                                                    }
                                                    loaderForObrSix.setFiveForFirstFormatTable(infoList, 6);
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    loaderForObrSix.setFiveForFirstFormatTable(infoList, 7);
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    loaderForObrSix.setFiveForFirstFormatTable(infoList, 8);
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    loaderForObrSix.setFourFormat(infoList, 9);
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    loaderForObrSix.setFourFormat(infoList, 10);
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    try {
                                                        if (infoList.phylum.size() != 0){
                                                            loaderForObrSix.setPhylum(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        }
                                                        loaderForObrSix.setSixForFirstFormatTable(infoList, 11);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        loaderForObrSix.setTwoFormatWithSer(infoList,15, "genus");
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        loaderForObrSix.setTwoFormatWithSer(infoList,16 , "species");;
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    } catch (XmlException e) {
                                                        e.printStackTrace();
                                                    } catch (IOException e) {
                                                        e.printStackTrace();
                                                    } catch (InvalidFormatException e) {
                                                        e.printStackTrace();
                                                    } catch (ClassNotFoundException e) {
                                                        e.printStackTrace();
                                                    }
                                                    loaderForObrSix.setAddition(infoList, 1);
                                                    loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                    try {
                                                        loaderForObrSix.saveFile(infoList, saveSampleDir);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        loaderForObrSix.getClose();
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        xlxsOpen.getClose();
                                                        loadStatus.setText("Загрузка: 100%");
                                                    } catch (IOException e) {
                                                        e.printStackTrace();
                                                    }
                                                    counter_files++;
                                                }
                                            }
                                            loadStatusFileNumber.setText("");
                                            loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                            checkStart = false;
                                        }
                                    }.start();
                                } else if (selectedSample.equals("Шаблон патогенная нагрузка")){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            loaderForObrSeven = new LoaderForObrSeven("obr_6");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getPhylum(infoList);
                                                            xlxsOpen.getGenus(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            xlxsOpen.getSpecies(infoList);
                                                            xlxsOpen.getBioIndex(infoList);
                                                            xlxsOpen.getPielouEveness(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        loaderForObrSeven.setFileNameForSecondFormatTable(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                        if (infoList.bioIndex.size() != 0){
                                                            loaderForObrSeven.setBioIndexForFirstTableFormat(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                        }
                                                        if (infoList.pielouEveness != null){
                                                            loaderForObrSeven.setPielouEveness(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                        }
                                                        try {
                                                            loaderForObrSeven.setSixForFirstFormatTable(infoList, 2);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            loaderForObrSeven.setTwoFormatWithSer(infoList, 4, "genus");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            loaderForObrSeven.setTwoFormatWithSer(infoList, 5, "species");
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                        } catch (XmlException e) {
                                                            e.printStackTrace();
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            loaderForObrSeven.saveFile(infoList, saveSampleDir);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            loaderForObrSeven.getClose();
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            xlxsOpen.getClose();
                                                            loadStatus.setText("Загрузка: 100%");
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                            }
                                        }.start();
                                    }
                                } else
                                {
                                    errorMessageStr = "Выбранная папка загрузки является пустой...";
                                    ErrorController errorController = new ErrorController();
                                    try {
                                        errorController.start(new Stage());
                                    } catch (IOException e) {
                                        e.printStackTrace();
                                    }
                                }
                            } else {
                                errorMessageStr = "Вы не выбрали шаблон для создания отчета...";
                                ErrorController errorController = new ErrorController();
                                try {
                                    errorController.start(new Stage());
                                } catch (IOException e) {
                                    e.printStackTrace();
                                }
                            }
                        } else {
                            errorMessageStr = "Вы не указаали директорию загрузки или директорию выгрузки...";
                            ErrorController errorController = new ErrorController();
                            try {
                                errorController.start(new Stage());
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    } else
                    {
                        errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
    }
}
