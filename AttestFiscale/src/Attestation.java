import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

public class Attestation {
	
	// Variables d'import des images 
	FileInputStream is;
	String imgLogo = "C:\\Users\\Jean\\Desktop\\Projetstage1ereannee\\DocGenerator\\Logo2.jpg";
	String imgSignature = "C:\\Users\\Jean\\Desktop\\Projetstage1ereannee\\DocGenerator\\Sanstitre.jpg";
	
	// Variable des informations de l'entreprise
	String entrepriseName = "Arkadia PC";
	String entrepriseStreet = "4, rue des Pyrénées";
	String entrepriseCity = "92500 Rueil Malmaison";
	String entreprisePhone = "+33 (1) 47 08 98 38";
	String entrepriseMail = "contact@arkadia-pc.fr";
	String entrepriseID = "Agrément N° SAP524160330";
	
	// Création de XWPFDocument
	XWPFDocument document = new XWPFDocument();
	
	// Création de la table
    XWPFTable table = document.createTable();
   
	// Appel de la  classe View 
 	View view;

 	// Méthode de création du Document
 	public void createDoc() {
 		try {
 			document = new XWPFDocument();
 			System.out.println("CREATED.");	
 		} catch (Exception e) {
 			e.printStackTrace();
 		}
 	}
 	
 	// Méthode de sauvegarde du Document
 	public void saveDoc(View view) {
 		try {
 			String file = "C:\\Users\\Jean\\Desktop\\Projetstage1ereannee\\DocGenerator\\Attestation-Fiscale-"+ view.getYearChooser() + "-" +view.getTxtPrenom() + "-" + view.getTxtNom() + ".doc";
 			FileOutputStream output = new FileOutputStream(file);
 			document.write(output);
 			output.close();
 			System.out.println("SAVED.");
 		} catch (Exception e) {
 			e.printStackTrace();
 		}
 	}
	
	// Trash variable "aller à la ligne" 
	String cheat = "                ";
	
	 /**
	 * Constructeur de la classe Attestation
	 * @param attestation
	 * @param view
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 */
	public Attestation(View view) throws InvalidFormatException, IOException {
		
		// Enlever les bordures de la table 
		table.removeBorders();
		// Custom des marges du document
		CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
		CTPageMar pageMar = sectPr.addNewPgMar();
		pageMar.setLeft(BigInteger.valueOf(1000L));
		pageMar.setTop(BigInteger.valueOf(550L));
		pageMar.setRight(BigInteger.valueOf(1000L));
		pageMar.setBottom(BigInteger.valueOf(550L));
		// Espacement entre les lignes
		//documentTitle.setSpacingBefore(100);
			
		// TODO trouver une méthode pour aller à la ligne sans utiliser le "cheat"
		// Création du header partie gauche, informations entreprise
		XWPFTableRow row = table.getRow(0);
		row.getCell(0).setText(entrepriseName + cheat + cheat + cheat);
		row.getCell(0).setText(entrepriseStreet + cheat + cheat);
		row.getCell(0).setText(entrepriseCity + cheat);
		row.getCell(0).setText(cheat + entreprisePhone);
		row.getCell(0).setText(cheat + entrepriseMail + cheat);
		row.getCell(0).setText(entrepriseID);
		row.addNewTableCell();
		row.addNewTableCell();	
		 
		
		// Import du logo
		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		paragraph = row.getCell(2).addParagraph();
		run = paragraph.createRun();	
		try {
			is = new FileInputStream(imgLogo);
			run.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgLogo, Units.toEMU(110), Units.toEMU(80));
			paragraph.setAlignment(ParagraphAlignment.RIGHT);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}
		
		// Header droit
		XWPFParagraph paragraph2 = document.createParagraph();
		XWPFRun run2 = paragraph2.createRun();
		run2.addBreak();
		run2.setText(view.getTxtNom() + " " + view.getTxtPrenom());
		run2.addBreak();
		run2.setText(view.getTxtAdresse());
		run2.addBreak();
		run2.setText(view.getTxtCP() + " " + view.getTxtVille());
		run2.addBreak();
		run2.setText("le " + view.getDateChooser());
		run2.addBreak();
		run2.addBreak();
		run2.setFontFamily("Calibri");
		paragraph2.setAlignment(ParagraphAlignment.RIGHT);
		
		// Titre 
		XWPFParagraph paragraph3 = document.createParagraph();
		XWPFRun run3 = paragraph3.createRun();
		run3.setText("Attestation destinée au Centre des Impôts");
		run3.setFontSize(20);
		run3.addBreak();
		run3.setFontFamily("Calibri");
		run3.setUnderline(UnderlinePatterns.SINGLE);
		paragraph3.setAlignment(ParagraphAlignment.CENTER);
		
		// Body
		XWPFParagraph paragraph4 = document.createParagraph();
		XWPFRun run4 = paragraph4.createRun();
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.setText("Je soussigné Monsieur Adelino Araujo gérant de l'organisme "
				+ "agréé Arkadia PC certifie que " + view.getCmbTitre() + " " + view.getTxtPrenom() + " " + view.getTxtNom() + " a bénéficié d'assistance informatique à domicile, service à la personne :");
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.addTab();
		run4.setText("Montant total des factures de " + view.getYearChooser() + " : " + view.getTxtMontantAttest() + " euros");
		run4.addBreak();
		run4.addTab();
		run4.addTab();
		run4.setText("Montant total payé en Cesu préfinancés : 0 euros");
		run4.addBreak();
		run4.addBreak();
		run4.setText("Intervenants : ");
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.addTab();
		run4.setText("Adelino Araujo");
		run4.addBreak();
		run4.addBreak();
		run4.setText("Prestations :");
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.setText("Les sommes perçues pour financer les services à la personne sont à déduire de la valeur indiquée précédemment.");
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.setText("La déclaration engage la responsabilité du seul contribuable");
		run4.addBreak();
		run4.addBreak();
		// écriture font size différent : 10
		XWPFParagraph paragraph04 = document.createParagraph();
		XWPFRun run04 = paragraph04.createRun();
		run04.setText("* Pour les personnes utilisant le Chèque Emploi Service Universel, seul le montant financé personnellement est déductible. "
				+ "Une attestation est délivrée par les établissements qui préfinancent le CESU.");
		run04.setFontSize(10);
		// retour au font size 11
		XWPFParagraph paragraph004 = document.createParagraph();
		XWPFRun run004 = paragraph004.createRun();
		run004.addBreak();
		run004.addBreak();
		run004.addTab();
		run004.setText("Fait pour valoir ce que de droit,");
		run004.addBreak();
		run004.addBreak(); 	
		run004.addBreak();
		run004.addBreak();
		run004.setText("Araujo Adelino, gérant.");
		run004.setFontFamily("Calibri");
		paragraph004.setIndentationLeft(0);
		paragraph004.setIndentationHanging(100);
		
		// import de la signature 
		XWPFParagraph paragraph5 = document.createParagraph();
		XWPFRun run5 = paragraph5.createRun();
		try {
			is = new FileInputStream(imgSignature);
			run5.addBreak();
			run5.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgSignature, Units.toEMU(200), Units.toEMU(70));
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		paragraph5.setAlignment(ParagraphAlignment.RIGHT);
		//paragraph4.setIndentationLeft(0);
		//paragraph4.setIndentationHanging(100);
		
		
	}
	
}
