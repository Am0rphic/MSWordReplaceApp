package application;
	
import javafx.application.Application;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.layout.BorderPane;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.stream.Stream;

public class MSWordReplace extends Application {
  private ArrayList<String> wordsToChange = new ArrayList<>(); 
  private ArrayList<String> templateTranslate = new ArrayList<>(); 
  private int counter; 
  private String filePath;
  private String userInput;
  
  @Override
  public void start(Stage app) {
     BorderPane layout = new BorderPane();
     TextField userInputText = new TextField("C:\\Users\\Am0rphic\\Downloads\\test.doc");
     Label labelText = new Label("Filename");
     Button nextButton = new Button("Next");
     userInput = "";
     layout.setTop(labelText);
     layout.setCenter(userInputText);
     layout.setBottom(nextButton);
     layout.setPrefSize(300,300);
     
     Scene testScene = new Scene(layout);
     app.setScene(testScene);
     app.show();
     
     nextButton.setOnAction((event)-> {
         if (labelText.getText().equals("Filename")) {
              filePath = userInputText.getText(); 
              readVariables(filePath, wordsToChange); 
              counter = wordsToChange.size(); 
              labelText.setText("Change next element - " + wordsToChange.get(wordsToChange.size()-counter));
              userInputText.setText("");
         }	else if (counter==0) {  
        	 nextButton.setText("Replace all?");
             POIFSFileSystem fs = null;
             try {
                 fs = new POIFSFileSystem(new FileInputStream(filePath));
                 HWPFDocument doc = new HWPFDocument(fs);
                 doc = replaceText(doc);
                 saveWord(filePath+"_test.doc", doc);
                 app.close();
             }
             catch(FileNotFoundException e){
                 e.printStackTrace();
             }
             catch(IOException e){
                 e.printStackTrace();
             }
         }    else if (counter>=1) {
             counter--; 
             System.out.println(counter);
             userInput = userInputText.getText();
             templateTranslate.add(userInput);
             userInputText.setText("");
        	 	if (counter>=1) {
        	 		labelText.setText("Change next element - " + wordsToChange.get(wordsToChange.size()-counter));
        	 	}	else {
        	 		labelText.setText("Replace all?");
        	 	}
         }
     });     
 }

	public static void main(String[] args) {
		launch(MSWordReplace.class);
	}
	
  private HWPFDocument replaceText(HWPFDocument doc){
      Range r1 = doc.getRange();

      for (int i = 0; i < r1.numSections(); ++i ) {
          Section s = r1.getSection(i);
          for (int x = 0; x < s.numParagraphs(); x++) {
              Paragraph p = s.getParagraph(x);
              for (int z = 0; z < p.numCharacterRuns(); z++) {
                  CharacterRun run = p.getCharacterRun(z);
                  String text = run.text();
                  for (int g=0; g<wordsToChange.size(); g++) {
                      if(text.contains(wordsToChange.get(g))) { 
                          run.replaceText(wordsToChange.get(g), templateTranslate.get(g));
                      }
                  }
              }
          }
      }
      return doc;
  }
  
  private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException {
      FileOutputStream out = null;
      try {
          out = new FileOutputStream(filePath);
          doc.write(out);
      } finally {
          out.close();
      }

  }

  private static void readVariables(String filePath, ArrayList<String> wordsToChange) { 
      POIFSFileSystem fs = null;
      try {
          fs = new POIFSFileSystem(new FileInputStream(filePath));
          HWPFDocument doc = new HWPFDocument(fs);
          String wholeDoc = doc.getDocumentText();
          String [] wholeDocWords = wholeDoc.split("[^a-zA-Z0-9$]");
          Stream<String> stream = Arrays.stream(wholeDocWords);
          stream.filter(var -> var.contains("$"))  
                  .forEach((word) -> {
                      if (!wordsToChange.contains(word)) { 
                          wordsToChange.add(word);
                      }
                  });
          doc.close();
      }
      catch(FileNotFoundException e){
          e.printStackTrace();
      }
      catch(IOException e){
          e.printStackTrace();
      }
  }
}