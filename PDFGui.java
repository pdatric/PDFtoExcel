package pdftext;

import javafx.application.Application;
import javafx.stage.Stage;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.scene.Scene;
import javafx.scene.layout.VBox;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.scene.paint.Color;
import javafx.scene.control.Label;
import javafx.scene.control.Button;
import javafx.geometry.Pos;
import javafx.geometry.Insets;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.scene.input.MouseEvent;
import pdftext.PDFTest;

public class PDFGui
        extends Application {

	private Text actionStatus;
	private Stage savedStage;
	private static final String titleTxt = "JavaFX File Chooser Example 1";

	public static void main(String [] args) {

		Application.launch(args);
	}

	@Override
	public void start(Stage primaryStage) {
	
		primaryStage.setTitle(titleTxt);	

		// Window label
		Label label = new Label("Select File Choosers");
		label.setTextFill(Color.DARKBLUE);
		label.setFont(Font.font("Calibri", FontWeight.BOLD, 36));
		HBox labelHb = new HBox();
		labelHb.setAlignment(Pos.CENTER);
		labelHb.getChildren().add(label);

		// Buttons
		
		Button btn2 = new Button("Choose multiple PDF files...");
		btn2.setOnAction(new MultipleFcButtonListener());
		HBox buttonHb2 = new HBox(10);
		buttonHb2.setAlignment(Pos.CENTER);
		buttonHb2.getChildren().addAll(btn2);
                
                

		// Status message text
		actionStatus = new Text();
		actionStatus.setFont(Font.font("Calibri", FontWeight.NORMAL, 20));
		actionStatus.setFill(Color.FIREBRICK);

		// Vbox
		VBox vbox = new VBox(30);
		vbox.setPadding(new Insets(25, 25, 25, 25));;
		vbox.getChildren().addAll(labelHb, buttonHb2, actionStatus);

		// Scene
		Scene scene = new Scene(vbox, 500, 300); // w x h
		primaryStage.setScene(scene);
		primaryStage.show();

		savedStage = primaryStage;
	}
        

	

	private class MultipleFcButtonListener implements EventHandler<ActionEvent> {

		@Override
		public void handle(ActionEvent e) {

                    try {
                        showMultipleFileChooser();
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(PDFGui.class.getName()).log(Level.SEVERE, null, ex);
                    }
		}
	}

	private List<File> showMultipleFileChooser() throws FileNotFoundException {

		FileChooser fileChooser = new FileChooser();
		fileChooser.setTitle("Select PDF files");
		
		fileChooser.getExtensionFilters().addAll(
			new ExtensionFilter("PDF Files", "*.pdf"));
		List<File> selectedFiles = fileChooser.showOpenMultipleDialog(savedStage);
                    
		if (selectedFiles != null) {
                        
			actionStatus.setText("PDF Files selected [" + selectedFiles.size() + "]: " +
					selectedFiles.get(0).getName() + "..");
                        int size = selectedFiles.size();
                        PDFTest.excelTemplate(selectedFiles, size);
                }
                return selectedFiles;
        }
                public static void go(String templatePath, List selectedFiles, int size){ 
                     
                if (selectedFiles !=null){     
                      for (int i = 0; i < selectedFiles.size(); i++){
                          String fPath = selectedFiles.get(i).toString();
                          System.out.println("LOOK AT ME " + fPath);
                          PDFTest.pdfTotxt(fPath, size, templatePath);
                      }  
                }
		/*else {
			actionStatus.setText("PDF file selection cancelled.");
		}*/
                
                
                
	}
        
       
}