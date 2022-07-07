import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import javax.swing.JFileChooser;
import javax.swing.JFrame;

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

public class AttestationModel {
	
	// Appel de la classe View 
	AttestationApplication attestationApplication;
	
	// Trash variable "aller à la ligne" 
	private static final String CHEAT = "                ";
	
	// Variable des informations de l'entreprise
	private static final String ENTREPRISE_HOLDER = "Adelino Araujo";
	private static final String ENTREPRISE_NAME = "Arkadia PC";
	private static final String ENTREPRISE_STREET = "4, rue des Pyrénées";
	private static final String ENTREPRISE_CITY = "92500 Rueil Malmaison";
	private static final String ENTREPRISE_PHONE = "+33 (1) 47 08 98 38";
	private static final String ENTREPRISE_MAIL = "contact@arkadia-pc.fr";
	private static final String ENTREPRISE_ID = "Agrément N° SAP524160330";
	
	// Variables d'import des images 
	private InputStream is;
	private File imgLogo = new File("src/img/logofinal.jpg");
	private String imgLogoAbsolute = imgLogo.getAbsolutePath();
	private File imgSignature = new File("src/img/signature.jpg");
	private String imgSignatureAbsolute = imgSignature.getAbsolutePath();
		
	
	// Création de XWPFDocument
	XWPFDocument document = new XWPFDocument();
	// Création de la table
    XWPFTable table = document.createTable();
   
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
 	public void saveDoc() throws IOException {
 		
 		// Constructeurs
 		JFrame parentFrame = new JFrame();
 		JFileChooser fileChooser = new JFileChooser();
 		
 		// Méthodes création de fichier
 		fileChooser.setDialogTitle("Enregistrer sous");
 		fileChooser.setSelectedFile(new File("Attestation-Fiscale" + attestationApplication.getYearChooser() + "-" + attestationApplication.getTxtPrenom() + "-" + attestationApplication.getTxtNom() + ".doc"));
 		int userSelection = fileChooser.showSaveDialog(parentFrame);
 		if (userSelection == JFileChooser.APPROVE_OPTION) {
 		   File fileToSave = fileChooser.getSelectedFile();
 		   FileOutputStream output;
 		   try {
 			   output = new FileOutputStream(fileToSave.getAbsolutePath());
 			   document.write(output);
 			   output.close();
 			   System.out.println("Sauvegarde du document: " + fileToSave);
 		   } catch (FileNotFoundException e) {
 			   e.printStackTrace();
 		   }
 		}
 	}

	
	 /**
	 * Constructeur de la classe Attestation
	 * @param attestation
	 * @param view
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 */
	public AttestationModel(AttestationApplication attestationApplication) throws InvalidFormatException, IOException {
		
		this.attestationApplication = attestationApplication;
		
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
		row.getCell(0).setText(ENTREPRISE_NAME + CHEAT + CHEAT + CHEAT);
		row.getCell(0).setText(ENTREPRISE_STREET + CHEAT + CHEAT);
		row.getCell(0).setText(ENTREPRISE_CITY + CHEAT);
		row.getCell(0).setText(CHEAT + ENTREPRISE_PHONE);
		row.getCell(0).setText(CHEAT + ENTREPRISE_MAIL + CHEAT);
		row.getCell(0).setText(ENTREPRISE_ID);
		row.addNewTableCell();
		row.addNewTableCell();	
		 
		
		// Import du logo
		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		paragraph = row.getCell(2).addParagraph();
		run = paragraph.createRun();	
		try {
			is = new FileInputStream(imgLogoAbsolute);
			run.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgLogoAbsolute, Units.toEMU(110), Units.toEMU(80));
			paragraph.setAlignment(ParagraphAlignment.RIGHT);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}
		
		// Header droit
		XWPFParagraph paragraph2 = document.createParagraph();
		XWPFRun run2 = paragraph2.createRun();
		run2.addBreak();
		run2.setText(attestationApplication.getTxtNom() + " " + attestationApplication.getTxtPrenom());
		run2.addBreak();
		run2.setText(attestationApplication.getTxtAdresse());
		run2.addBreak();
		run2.setText(attestationApplication.getTxtCP() + " " + attestationApplication.getTxtVille());
		run2.addBreak();
		run2.addBreak();
		run2.setText("le " + attestationApplication.getDateChooser() + ",");
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
		run4.setText("Je soussigné Monsieur " + ENTREPRISE_HOLDER +  " gérant de l'organisme agréé " + ENTREPRISE_NAME + " certifie que " + attestationApplication.getCmbTitre() + " " + attestationApplication.getTxtPrenom() + " " + attestationApplication.getTxtNom() + " a bénéficié d'assistance informatique à domicile, service à la personne :");
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.addTab();
		run4.setText("Montant total des factures de " + attestationApplication.getYearChooser() + " : " + attestationApplication.getTxtMontantAttest() + " euros");
		run4.addBreak();
		run4.addTab();
		run4.addTab();
		run4.setText("Montant total payé en CESU préfinancé : 0 euros"); // Cesu préfinancés ? 
		run4.addBreak();
		run4.addBreak();
		run4.setText("Intervenants : ");
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.addTab();
		run4.setText(ENTREPRISE_HOLDER);
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
		run004.setText(ENTREPRISE_HOLDER + ", gérant.");
		run004.setFontFamily("Calibri");
		paragraph004.setIndentationLeft(0);
		paragraph004.setIndentationHanging(100);
		
		// import de la signature 
		XWPFParagraph paragraph5 = document.createParagraph();
		XWPFRun run5 = paragraph5.createRun();
		try {
			is = new FileInputStream(imgSignatureAbsolute);
			run5.addBreak();
			run5.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgSignatureAbsolute, Units.toEMU(200), Units.toEMU(70));
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		paragraph5.setAlignment(ParagraphAlignment.RIGHT);
	}
	
}
