package com.example.ExcelAutomator;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.Pane;
import javafx.scene.paint.Color;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.sql.SQLException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.ResourceBundle;
import java.util.logging.Logger;

public class SceneController implements Initializable {

    @FXML
    public ProgressBar progressBar;
    public Text errorMsg;
    // Home Pane Elements
    @FXML
    private Pane home, sub_incentive_pane, emp_incentive_pane;
    @FXML
    private ScrollPane homeBackground;

    @FXML
    private TreeView selectionTree;
    // -------------------------------------------------------
    // MKS Employee Incentive Pane Elements
    @FXML
    private RadioButton emp_filing, emp_setting, emp_polishing;

    @FXML
    private DatePicker emp_inc_start, emp_inc_end;

    private LocalDate emp_start, emp_end; //?
    // --------------------------------------------------------
    // Subcontractor Incentive Pane Elements
    @FXML
    private RadioButton sub_filing, sub_setting, sub_polishing;
    @FXML
    private DatePicker sub_calendar_start, sub_calendar_end;
    @FXML
    private TextField sub_gold_rate_field, sub_plat_rate_field;
    @FXML
    private Button export_btn;
    //---------------------------------------------------------
    // Export Button




    String sub_incentive_type;
    String emp_incentive_type;
    int state;
    // 0 = home
    // 1 = employee incentive
    // 2 = sub incentive
    LocalDate sub_start;
    LocalDate sub_end;

    DirectoryChooser directoryChooser;

    int sub_gold_rate, sub_plat_rate;
    String path;

    ArrayList<XSSFWorkbook> workbooks;
    private static final Logger LOGGER = Logger.getLogger(SQLUtils.class.getName());

    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {

        // Home Pane Elements
        progressBar.setVisible(false);
        progressBar.setStyle("-fx-accent: #DF6D24");

        errorMsg.setVisible(false);
        errorMsg.setText("* Please add missing values *");
        errorMsg.setFill(Color.RED);

        directoryChooser = new DirectoryChooser();



        TreeItem<String> dummyRoot = new TreeItem<>();
        selectionTree.setRoot(dummyRoot);
        selectionTree.setShowRoot(false);

        TreeItem<String> acct = new TreeItem<>("Accounting");
        TreeItem<String> incentive = new TreeItem<>("Incentive Reports");
        dummyRoot.getChildren().addAll(acct, incentive);

        TreeItem<String> ar = new TreeItem<>("Accounts Receivable");
        TreeItem<String> ap = new TreeItem<>("Accounts Payable");
        TreeItem<String> mks = new TreeItem<>("MKS Employees");
        TreeItem<String> sub = new TreeItem<>("Subcontractors");

        acct.getChildren().addAll(ar, ap);
        incentive.getChildren().addAll(mks, sub);

        homeBackground = new ScrollPane();
        homeBackground.setFitToHeight(true);
        homeBackground.setFitToWidth(true);
        //---------------------------------------------------------
        // MKS Emp. Incentive Pane Elements
        ToggleGroup emp_incentive_select = new ToggleGroup();
        emp_filing.setToggleGroup(emp_incentive_select);
        emp_setting.setToggleGroup(emp_incentive_select);
        emp_polishing.setToggleGroup(emp_incentive_select);
        // need to add text boxes etc.
        //---------------------------------------------------------
        // Subcontractor Incentive Pane Elements
        ToggleGroup sub_incentive_select = new ToggleGroup();
        sub_filing.setToggleGroup(sub_incentive_select);
        sub_setting.setToggleGroup(sub_incentive_select);
        sub_polishing.setToggleGroup(sub_incentive_select);
        //---------------------------------------------------------

    }

    // visible pane controller: switch statement takes input from treeview selection and changes visible pane accordingly
    // selected pane visibility is set to true, all others are set to false
    public void selectPane(MouseEvent mouseEvent) {
        Object selection = selectionTree.getSelectionModel().getSelectedItem();
        System.out.println(selection);

        if (selection != null) {
            String menuChoice = selection.toString();

            switch (menuChoice) {
                case "TreeItem [ value: Accounts Receivable ]":
                    this.home.setVisible(true);
                    this.emp_incentive_pane.setVisible(false);
                    this.sub_incentive_pane.setVisible(false);
                    state = 0;
                    break;

                case "TreeItem [ value: MKS Employees ]":
                    this.home.setVisible(false);
                    this.emp_incentive_pane.setVisible(true);
                    this.sub_incentive_pane.setVisible(false);
                    state = 1;
                    break;

                case "TreeItem [ value: Subcontractors ]":
                    this.home.setVisible(false);
                    this.sub_incentive_pane.setVisible(true);
                    this.emp_incentive_pane.setVisible(false);
                    state = 2;
                    break;
            }
        }
    }

    // Emp. Incentive Pane - Data Extraction Logic
    public String get_emp_incentive_type(ActionEvent actionEvent){
        if (emp_filing.isSelected()){
            emp_incentive_type = emp_filing.getText();
        }
        if (emp_setting.isSelected()){
            emp_incentive_type = emp_setting.getText();
        }
        if (emp_polishing.isSelected()){
            emp_incentive_type = emp_polishing.getText();
        }
        return emp_incentive_type;
    }
    public void emp_incentive_period(ActionEvent actionEvent) {
    }
    //--------------------------------------------------------------
    // Subcontract Inventive Pane - Data Extraction Logic
    public void get_sub_incentive_type(ActionEvent actionEvent) {
        if (sub_filing.isSelected()) {
            sub_incentive_type = sub_filing.getText();
        }
        if (sub_polishing.isSelected()) {
            sub_incentive_type = sub_polishing.getText();
        }
        if (sub_setting.isSelected()) {
            sub_incentive_type = sub_setting.getText();
        }
    }

    public void sub_incentive_start(ActionEvent actionEvent) {
        sub_start = sub_calendar_start.getValue();
    }

    public void sub_incentive_end(ActionEvent actionEvent){
        sub_end = sub_calendar_end.getValue();
    }



    public void get_sub_gold_rate(ActionEvent actionEvent) {
//        sub_gold_rate = Integer.parseInt(sub_gold_rate_field.getText());
//        System.out.println(sub_gold_rate);
    }

    public void get_sub_platinum_rate(ActionEvent actionEvent) {
//        sub_plat_rate = Integer.parseInt(sub_plat_rate_field.getText());
//        System.out.println(sub_plat_rate);
    }


    //---------------------------------------------------------------
    // Export Logic

    public void handle_export(ActionEvent actionEvent) throws SQLException, IOException {
        if(!sub_incentive_pane.isVisible()){
            return;
        }
        if (this.sub_incentive_type == null || sub_gold_rate_field.getText().equals("") ||
                sub_plat_rate_field.getText().equals("") ||  sub_calendar_end.getValue() == null || sub_calendar_end.getValue() == null ){

            errorMsg.setVisible(true);
            return;
        }

        errorMsg.setVisible(false);

        switch (state){
            case 0: //home page
                //pop up saying plz select report or nothing
                break;
            case 1: //employee incentive
                // handle emp report selection, period, and metal rates etc.
                break;
            case 2: //sub incentive

                if (sub_incentive_type.equals("Filing")){
                    System.out.println("Generate filing report");
                } else if (sub_incentive_type.equals("Setting")){
                    System.out.println("generate setting report");
                    sub_start = sub_calendar_start.getValue();
                    sub_end = sub_calendar_end.getValue();
                    sub_gold_rate = Integer.parseInt(sub_gold_rate_field.getText());
                    System.out.println(sub_gold_rate);
                    sub_plat_rate = Integer.parseInt(sub_plat_rate_field.getText());


                    File file = directoryChooser.showDialog(new Stage());
                    if (file != null){
                        path = file.getAbsolutePath();
                    }
                    if (path == null){
                        return;
                    }
                    System.out.println("saved path");
                    System.out.println(path);



                    Thread retrieve_and_parse = new Thread(new Runnable() {
                        @Override
                        public void run() {
                            MKSConnect mksConnect = new MKSConnect();
                            ExportSubIncentive exportSubIncentive = new ExportSubIncentive();
                            progressBar.setVisible(true);
                            workbooks = new ArrayList<>();
                            workbooks = mksConnect.handleQuery(sub_start, sub_end,sub_gold_rate,sub_plat_rate,path);
                            System.out.println("workbooks returned to scene controller");
                            exportSubIncentive.formatSubInc(workbooks);
                            exportSubIncentive.exportWorkbooks(workbooks,path);
                            progressBar.setVisible(false);
                        }
                    });
                    retrieve_and_parse.start();

                } else if (sub_incentive_type.equals("Polishing")){
                    System.out.println("Generate polishing report");
                }
        }
    }




}
