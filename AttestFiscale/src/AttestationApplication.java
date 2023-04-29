import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.io.IOException;
import java.text.DateFormatSymbols;
import java.util.Calendar;

import javax.swing.AbstractAction;
import javax.swing.ActionMap;
import javax.swing.DefaultComboBoxModel;
import javax.swing.InputMap;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.border.EmptyBorder;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.toedter.calendar.JDateChooser;

/**
 * 
 * Classe principale de l'application
 * Hérite de la classe JFrame pour créer une fenêtre graphique
 */
public class AttestationApplication extends JFrame {
	
	
	/**
	* Variables du JFrame
	*/
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	private JTextField txtNom;
	private JTextField txtPrenom;
	private JTextField txtAdresse;
	private JTextField txtVille;
	private JTextField txtCP;
	private JTextField txtMontantAttest;
	private JComboBox<String> cmbTitre;
	
	/**
	 * 
     * Renvoie le nom saisi dans le champ de texte correspondant.
     * @return Le nom saisi dans le champ de texte.
     */
	public String getTxtNom() {
		return txtNom.getText();
	}

	/**
	 * 
	 * Renvoie le prénom saisi dans le champ de texte correspondant.
	 * @return Le prénom saisi dans le champ de texte.
     */
	public String getTxtPrenom() {
		return txtPrenom.getText();
	}
	
	/**
	 * 
     * Renvoie l'adresse saisie dans le champ de texte correspondant.
     * @return L'adresse saisie dans le champ de texte.
     */
	public String getTxtAdresse() {
		return txtAdresse.getText();
	}

	/**
	 * 
     * Renvoie la ville saisie dans le champ de texte correspondant.
     * @return La ville saisie dans le champ de texte.
     */
	public String getTxtVille() {
		return txtVille.getText();
	}

	/**
	 * 
 	 * Renvoie le code postal saisi dans le champ de texte correspondant.
     * @return Le code postal saisi dans le champ de texte.
     */
	public String getTxtCP() {
		return txtCP.getText();
	}

	/**
	 * 
 	 * Renvoie le montant de l'attestation saisi dans le champ de texte correspondant.
     * @return Le montant de l'attestation saisi dans le champ de texte.
     */
	public String getTxtMontantAttest() {
		return txtMontantAttest.getText();
	}

	/**
	 * 
	 * Renvoie le titre sélectionné dans la liste déroulante correspondante.
	 * @return Le titre sélectionné dans la liste déroulante. Si "Aucun titre" est sélectionné, une chaîne vide est renvoyée.
	 */
	public String getCmbTitre() {
		if (cmbTitre.getSelectedItem() == "Aucun titre") {
			return "";
		}
		return (String) cmbTitre.getSelectedItem();
	}
	
	/**
	 * 
	 * Méthode getDateChooser() implémentée de la méthode getMontForInt()
	 * @return la date jour int, mois letters, année int
	 */
	public String getDateChooser() {
		Calendar calendar = Calendar.getInstance();
		return calendar.get(Calendar.DAY_OF_MONTH) + 
				" " + getMonthForInt(calendar.get(Calendar.MONTH)) + 
				" " + calendar.get(Calendar.YEAR);
	}
	
	/**
	 * 
	 * Retourne le nom du mois correspondant au nombre entier donné en paramètre.
	 * @param m le nombre entier représentant le mois (entre 0 et 11, inclus).
	 * @return le nom du mois correspondant, sous forme de chaîne de caractères.
	 */
	String getMonthForInt(int m) {
	    String month = "";
	    DateFormatSymbols dfs = new DateFormatSymbols();
	    String[] months = dfs.getMonths();
	    if (m >= 0 && m <= 11 ) {
	        month = months[m];
	    }
	    return month;
	}
	
	/**
	 * 
	 * Méthode sélection de l'année uniquement
	 * @return l'année indiquée dans le calendrier
	 */
	public int getYearChooser() {
		Calendar calendar = Calendar.getInstance();
		return calendar.get(Calendar.YEAR);
	}
	
	/**
	 * 
	 * Méthode principale de lancement de l'application
	 * @param args tableau de chaînes de caractères représentant les arguments de la ligne de commande
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					AttestationApplication frame = new AttestationApplication();
					AttestationModel attestationModel = new AttestationModel(frame);
					attestationModel.createDoc();
					frame.setLocationRelativeTo(null); // centrer l'application
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}
	
	/**
	 * 
	 * Vérifie que tous les champs obligatoires sont remplis avant d'appeler la méthode "save" pour générer l'attestation fiscale.
     * Si tous les champs sont remplis, appelle la méthode "save".
     * @throws InvalidFormatException si le format d'un champ est invalide.
     * @throws IOException en cas d'erreur d'entrée/sortie lors de la création du document.
    */
	public void isInputValid() throws InvalidFormatException, IOException {
		if(("".equals(getTxtNom())) || "".equals(getTxtPrenom()) || "".equals(getTxtVille()) || 
				"".equals(getTxtAdresse()) || "".equals(getTxtMontantAttest())) {
			JOptionPane.showMessageDialog(contentPane, "Merci de remplir tous les champs");
		} else {
			try {
				save();
				JOptionPane.showMessageDialog(contentPane, "Sauvegarde réussie !");			
			} catch (Exception ex) {
				ex.printStackTrace();
		        JOptionPane.showMessageDialog(contentPane, "Erreur lors de la sauvegarde.");
			}
		}
	}
	
	/**
	 * 
	 * Méthode de fermeture de l'application. Affiche une boîte de dialogue de confirmation de fermeture.
     * Si l'utilisateur clique sur "Oui", l'application se ferme.
     */
	public void close() {
		int n = JOptionPane.showOptionDialog(new JFrame(), "Fermer application?", "Quitter", 
				JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, new Object[] {"Oui", "Non"}, 
				JOptionPane.YES_OPTION);
		if (n == JOptionPane.YES_OPTION) {
			dispose();
		} 	
	}
	
	/**  
	 * 
	 * Méthode de sauvegarde du document
     * Appelle la méthode saveDoc()
     * @throws IOException si une erreur d'entrée/sortie se produit
     * @throws InvalidFormatException si le format de fichier est invalide
     */
	public void save() throws IOException, InvalidFormatException {
		AttestationModel attestation = new AttestationModel(this);
		int n = JOptionPane.showOptionDialog(new JFrame(), "Confirmer enregistrement", "Enregistrer", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, new Object[] {"Oui", "Non"}, JOptionPane.YES_OPTION);
		if (n == JOptionPane.YES_OPTION) {
			attestation.saveDoc();
		} 
	}

	/**
	 * Création du Frame 
	 */
	public AttestationApplication() {
		
		/**
		 * Information de création du JFrame
		 */
		setTitle("Attestation Fiscale");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 457, 501);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		/**
		 * Titre
		 */
		JLabel lblTitre = new JLabel("Titre");
		lblTitre.setBounds(34, 36, 46, 14);
		contentPane.add(lblTitre);
	
		cmbTitre = new JComboBox<String>();
		cmbTitre.setModel(new DefaultComboBoxModel<String>(new String[] {"Madame", "Mademoiselle", "Monsieur", "Aucun titre"}));
		cmbTitre.setBounds(34, 61, 153, 22);
		contentPane.add(cmbTitre);
		
		/**
		 * Nom
		 */
		JLabel lblNom = new JLabel("Nom");
		lblNom.setBounds(34, 100, 46, 14);
		contentPane.add(lblNom);
		
		txtNom = new JTextField();
		txtNom.setBounds(34, 125, 153, 20);
		contentPane.add(txtNom);
		txtNom.setColumns(10);
		
		/**
		 * Prénom
		 */
		JLabel lblPrenom = new JLabel("Prénom");
		lblPrenom.setBounds(240, 100, 46, 14);
		contentPane.add(lblPrenom);
		
		txtPrenom = new JTextField();
		txtPrenom.setColumns(10);
		txtPrenom.setBounds(240, 125, 153, 20);
		contentPane.add(txtPrenom);
		
		/**
		 * Adresse
		 */
		JLabel lblAdresse = new JLabel("Adresse");
		lblAdresse.setBounds(34, 156, 153, 14);
		contentPane.add(lblAdresse);
		
		txtAdresse = new JTextField();
		txtAdresse.setColumns(10);
		txtAdresse.setBounds(34, 178, 359, 20);
		contentPane.add(txtAdresse);
		
		/**
		 * Ville
		 */
		JLabel lblVille = new JLabel("Ville");
		lblVille.setBounds(34, 209, 46, 14);
		contentPane.add(lblVille);
		
		txtVille = new JTextField();
		txtVille.setColumns(10);
		txtVille.setBounds(34, 234, 216, 20);
		contentPane.add(txtVille);
		
		/**
		 * Code Postal
		 */
		JLabel lblCP = new JLabel("Code Postal");
		lblCP.setBounds(298, 209, 95, 14);
		contentPane.add(lblCP);
		
		txtCP = new JTextField();
		txtCP.setColumns(10);
		txtCP.setBounds(298, 234, 95, 20);
		contentPane.add(txtCP);
		
		/**
		 * Montant attestation
		 */
		JLabel lblMontantAttest = new JLabel("Montant attestation");
		lblMontantAttest.setBounds(34, 285, 115, 14);
		contentPane.add(lblMontantAttest);
		
		txtMontantAttest = new JTextField();
		txtMontantAttest.setColumns(10);
		txtMontantAttest.setBounds(34, 310, 153, 20);
		contentPane.add(txtMontantAttest);
		

		/**
		 * Date attestation
		 */
		JLabel lblDate = new JLabel("Date attestation");
		lblDate.setBounds(298, 285, 95, 14);
		contentPane.add(lblDate);
		
		JDateChooser dateChooser = new JDateChooser(); 
		dateChooser.setDateFormatString("dd MMMM yyyy");
		dateChooser.setCalendar(Calendar.getInstance()); // set la date du jour dans le frame
		dateChooser.setBounds(240, 310, 153, 20);
		contentPane.add(dateChooser);
		//dateChooser.setEnabled(false);
		
		/**
		 * 
 		 * Bouton enregistrer pour sauvegarder les informations dans le document
	     * Ajoute les événements de clic et de touche pour appeler la méthode isInputValid()
	     */
		JButton btnEnregistrer = new JButton("Enregistrer");
		// Appel de la méthode isInputValid() lorsqu'on appuie sur la touche Enter
		btnEnregistrer.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				try {
					if(e.getKeyCode() == KeyEvent.VK_ENTER) {
					isInputValid();	
					}
				} catch (InvalidFormatException | IOException e1) {
					e1.printStackTrace();
				}
			}
		});
		// Appel de la méthode isInputValid() lorsqu'on clique sur le bouton enregistrer
		btnEnregistrer.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					isInputValid();
				} catch (InvalidFormatException e1) {
					e1.printStackTrace();
				} catch (IOException ex) {
					ex.printStackTrace();
				}
			}
		});
		btnEnregistrer.setFont(new Font("Tahoma", Font.BOLD, 14));
		btnEnregistrer.setBounds(34, 370, 153, 48);
		contentPane.add(btnEnregistrer);
		
		/**
		 * 
		 * Bouton quitter
		 * Gère les événements clic bouton et touche échappement
	     */
		JButton btnQuitter = new JButton("Quitter");
		// Appelle la méthode close() lors de l'événement touche échappement
		InputMap im = getRootPane().getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW);
		ActionMap am = getRootPane().getActionMap();
		im.put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0), "cancel");
		am.put("cancel", new AbstractAction() {
			private static final long serialVersionUID = 1L;
			@Override
			public void actionPerformed(ActionEvent e) {
				close();
			}
		});
		// Appelle la méthode close() lors de l'événement clic bouton quitter
		btnQuitter.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				close();
			}
		});
		btnQuitter.setFont(new Font("Tahoma", Font.BOLD, 14));
		btnQuitter.setBounds(240, 370, 153, 48);
		contentPane.add(btnQuitter);
	}
}
