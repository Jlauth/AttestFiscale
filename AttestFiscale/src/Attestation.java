import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class Attestation {
	
	// Création de XWPFDocument
	XWPFDocument document = new XWPFDocument();
	
	// variable d'import de l'image
	FileInputStream is;
	
	// appel de la  classe View 
	View view;
	
	/**
	 * Constructeur de la classe Attestation
	 * @param attestation
	 * @param view
	 */
	public Attestation(View view) {
	
		
		// import du logo
		XWPFParagraph paragraph11 = document.createParagraph();
		XWPFRun run11 = paragraph11.createRun();
		String imgLogo = "C:\\Users\\Jean\\Desktop\\Projetstage1ereannee\\DocGenerator\\Logo2.jpg";
		try {
			is = new FileInputStream(imgLogo);
			run11.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgLogo, Units.toEMU(100), Units.toEMU(75));
			paragraph11.setAlignment(ParagraphAlignment.RIGHT);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		// header partie gauche, Arkadia PC bold
		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setText("Arkadia PC");
		run.addBreak();
		run.setBold(true);
		run.setFontFamily("Calibri");
		paragraph.setAlignment(ParagraphAlignment.LEFT);
		
		// header partie gauche, informations Arkadia PC
		XWPFParagraph paragraph1 = document.createParagraph();
		XWPFRun run1 = paragraph.createRun();
		run1.setText("4, rue des Pyrénées");
		run1.addBreak();
		run1.setText("92500 Rueil Malmaison");
		run1.addBreak();
		run1.setText("+33 (1) 47 08 98 38");
		run1.addBreak();
		run1.setText("contact@arkadia-pc.fr");
		run1.addBreak();
		run1.addBreak();
		run1.setText("Agrément N° SAP524160330");
		run1.setFontFamily("Calibri");
		paragraph1.setAlignment(ParagraphAlignment.LEFT);
		
		// header partie droite
		XWPFParagraph paragraph2 = document.createParagraph();
		XWPFRun run2 = paragraph2.createRun();
		run2.setText(view.getTxtNom() + " " + view.getTxtPrenom());
		run2.addBreak();
		run2.setText(view.getTxtAdresse());
		run2.addBreak();
		run2.setText(view.getTxtCP() + " " + view.getTxtVille());
		run2.addBreak();
		run2.setText("le " + view.getDateChooser() + " .");
		run2.setFontFamily("Calibri");
		paragraph2.setAlignment(ParagraphAlignment.RIGHT);
		
		// titre 
		XWPFParagraph paragraph3 = document.createParagraph();
		XWPFRun run3 = paragraph3.createRun();
		run3.setText("Attestation destinée au Centre des Impôts");
		run3.setFontSize(20);
		run3.addBreak();
		run3.setFontFamily("Calibri");
		run3.setUnderline(UnderlinePatterns.SINGLE);
		paragraph3.setAlignment(ParagraphAlignment.CENTER);
		
		// body
		XWPFParagraph paragraph4 = document.createParagraph();
		XWPFRun run4 = paragraph4.createRun();
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.setText("Je soussigné Monsieur Adelino Araujo gérant de l'organisme "
				+ "agréé Arkadia PC certifie que Monsieur " + view.getTxtPrenom() + " " + view.getTxtNom() + " a bénéficié d'assistance informatique à domicile, service à la personne :");
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.addTab();
		run4.setText("Montant total des factures de " + view.getDateChooser() + " : " + view.getTxtMontantAttest());
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
		run4.setText("* Pour les personnes utilisant le Chèque Emploi Service Universel, seul le montant financé personnellement est déductible. "
				+ "Une attestation est délivrée par les établissements qui préfinancent le CESU.");
		run4.addBreak();
		run4.addBreak();
		run4.addBreak();
		run4.addTab();
		run4.setText("Fait pour valoir ce que de droit,");
		run4.addBreak();
		run4.addBreak(); 	
		run4.setText("Araujo Adelino, gérant.");
		run4.setFontFamily("Calibri");
		paragraph4.setIndentationLeft(0);
		paragraph4.setIndentationHanging(100);
		
		// import de la signature 
		XWPFParagraph paragraph5 = document.createParagraph();
		XWPFRun run5 = paragraph5.createRun();
		String imgSignature = "C:\\Users\\Jean\\Desktop\\Projetstage1ereannee\\DocGenerator\\Sanstitre.jpg";
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
		

	// création du Document
	public void createDoc() {
		try {
			document = new XWPFDocument();
			System.out.println("CREATED.");	
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	// sauvegarde du Document
	public void saveDoc() {
		try {
			FileOutputStream output = new FileOutputStream("C:\\Users\\Jean\\Desktop\\Projetstage1ereannee\\DocGenerator\\attestation-fiscale-annee-prenom-nom.doc");
			document.write(output);
			output.close();
			System.out.println("SAVED.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	
}
