import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;

import javax.imageio.ImageIO;
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

/**
* Cette classe représente le modèle d'une attestation fiscale qui peut être utilisée pour générer un document Word.
* Elle utilise l'application d'attestation fiscale pour extraire les informations nécessaires pour créer l'attestation.
* La classe crée un document Word et ajoute les informations de l'attestation dans un tableau pré-formaté.
* La classe utilise Apache POI pour créer et manipuler le document Word.
*/
public class AttestationModel {
	
	/**
	 * Déclaration de la variable pour l'instance de l'application Attestation
     */
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
	
	// Chemin relatif des images par rapport au package courant
	String logoPath = "/img/logofinal.jpg";
	String signaturePath = "/img/signature.jpg";

	// Récupération des InputStream correspondants aux images
	InputStream logoInputStream = getClass().getResourceAsStream(logoPath);
	InputStream signatureInputStream = getClass().getResourceAsStream(signaturePath);

	// Utilisation des InputStream pour charger les images
	// BufferedImage logoImage = ImageIO.read(logoInputStream);
	// BufferedImage signatureImage = ImageIO.read(signatureInputStream);
		
	/**
	 * 
	 * Crée un nouveau document XWPF et une table
	 */
	XWPFDocument document = new XWPFDocument();
    XWPFTable table = document.createTable();
   
    /**
     * Méthode de création de document Word.
     * Crée un nouveau document XWPFDocument.
     */
 	public void createDoc() {
 		try {
 			document = new XWPFDocument();
 			System.out.println("CREATED.");	
 		} catch (Exception e) {
 			e.printStackTrace();
 		}
 	}
 	
 	/**
 	 * Cette méthode permet de sauvegarder un document Word en utilisant un JFileChooser pour spécifier l'emplacement et le nom du fichier.
 	 * Elle crée le nom de fichier à partir des informations saisies dans l'application.
 	 * Une fois le fichier sélectionné et validé par l'utilisateur, le document est sauvegardé dans l'emplacement spécifié.
 	 * 
 	 * @throws IOException si une erreur d'entrée/sortie se produit lors de la sauvegarde du document.
 	 */
 	public void saveDoc() throws IOException {
 		// Constructeurs du frame pour la sauvegarde
 		JFrame parentFrame = new JFrame();
 		JFileChooser fileChooser = new JFileChooser();
 		// Méthodes de création du fichier d'enregistrement
 		fileChooser.setDialogTitle("Enregistrer sous");
 		fileChooser.setSelectedFile(new File("Attestation-Fiscale" + attestationApplication.getYearChooser() + 
 				"-" + attestationApplication.getTxtPrenom() + "-" + attestationApplication.getTxtNom() + ".doc"));
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
    *Constructeur de la classe AttestationModel qui initialise le modèle d'une attestation fiscale.
	*
	*@param attestationApplication l'application d'attestation fiscale à partir de laquelle les informations sont extraites pour créer l'attestation.
	*@throws InvalidFormatException si le format du document Word est invalide.
	*@throws IOException si une erreur d'entrée/sortie se produit lors de la création du document ou de l'ajout du logo.
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
			InputStream is = getClass().getResourceAsStream("/img/logofinal.jpg");
			BufferedImage logoImage = ImageIO.read(is);
			ByteArrayOutputStream os = new ByteArrayOutputStream();
			ImageIO.write(logoImage, "jpeg", os);
			InputStream logoInputStream = new ByteArrayInputStream(os.toByteArray());
			run.addPicture(logoInputStream, XWPFDocument.PICTURE_TYPE_JPEG, "logofinal.jpg", Units.toEMU(110), Units.toEMU(80));
			paragraph.setAlignment(ParagraphAlignment.RIGHT);
		} catch (IOException e1) {
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
		    InputStream is = getClass().getResourceAsStream("/img/signature.jpg");
		    run5.addBreak();
		    run5.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, "signature.jpg", Units.toEMU(200), Units.toEMU(70));
		} catch (InvalidFormatException e) {
		    e.printStackTrace();
		} catch (IOException e) {
		    e.printStackTrace();
		}
		paragraph5.setAlignment(ParagraphAlignment.RIGHT);
	}
}
