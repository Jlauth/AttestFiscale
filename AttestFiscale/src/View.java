import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.KeyStroke;

import com.toedter.calendar.JDateChooser;
import javax.swing.JButton;
import java.awt.Font;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.AbstractAction;
import javax.swing.ActionMap;
import javax.swing.DefaultComboBoxModel;
import javax.swing.InputMap;

import java.awt.event.ActionListener;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.awt.event.ActionEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;

public class View extends JFrame {

	private JPanel contentPane;
	private JTextField txtNom;
	private JTextField txtPrenom;
	private JTextField txtAdresse;
	private JTextField txtVille;
	private JTextField txtCP;
	private JTextField txtMontantAttest;
	private JComboBox cmbtitre;
	private JDateChooser dateChooser;
	
	/**
	 * méthode main de lancement de l'application
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					View frame = new View();
					Attestation attest = new Attestation(frame);
					attest.createDoc();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}
	
	/*
	public void checkFields() {
		if (validateFields()) {
			txtNom.getText();
			txtPrenom.getText();
			txtAdresse.getText();
			txtVille.getText();
			txtCP.getText();
			txtMontantAttest.getText(); 
			}
	}

	public boolean validateFields() {
		if (!validateField(txtNom, "Indiquer un nom")) {
			return false;
		} else if (!validateField(txtPrenom, "Indiquer un prénom")) {
			return false;
		} else if (!validateField(txtAdresse, "Indiquer un texte")) {
			return false;
		} else if (!validateField(txtVille, "Indiquer une ville")) {
			return false;
		} else if (!validateField(txtCP, "Indiquer un code postal")) {
			return false;
		} else if (!validateField(txtMontantAttest, "Indiquer un montant")) {
			return false;
		}
		return true;
	}
	
	public boolean validateField(JTextField field, String errormsg) {
		if(field.getText().equals("")) {
			return failedMessage(field, errormsg); 
		}
		else {
			return true;
		}	
	}
	*/
	
	/**
	 * Vérification champs remplis
	 * Appel de la méthode save si ils sont remplis
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public void isInputValid() throws InvalidFormatException, IOException {
		if(("".equals(getTxtNom())) || "".equals(getTxtPrenom()) || "".equals(getTxtVille()) || "".equals(getTxtAdresse()) || "".equals(getTxtMontantAttest())) {
			JOptionPane.showMessageDialog(contentPane, "Merci de remplir tous les champs");
		} else {
			save();
		}	
	}
	
	/**
	 * Méthode fermeture de l'application
	 */
	public void close() {
		int n = JOptionPane.showOptionDialog(new JFrame(), "Fermer application?", "Quitter", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, new Object[] {"Oui", "Non"}, JOptionPane.YES_OPTION);
		if (n == JOptionPane.YES_OPTION) {
			dispose();
		} 	
	}
	
	/**
	 * Méthode bouton enregistrer 
	 * Appel de saveDoc();
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 */
	public void save() throws InvalidFormatException, IOException {
		Attestation attestation = new Attestation(this);
		int n = JOptionPane.showOptionDialog(new JFrame(), "Confirmer enregistrement", "Enregistrer", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, new Object[] {"Oui", "Non"}, JOptionPane.YES_OPTION);
		if (n == JOptionPane.YES_OPTION) {
			attestation.saveDoc();
		} 
	}
	
	/**
	 * Getters & Setters
	 */
	public String getTxtNom() {
		return txtNom.getText();
	}

	public String getTxtPrenom() {
		return txtPrenom.getText();
	}
	
	public String getTxtAdresse() {
		return txtAdresse.getText();
	}

	public String getTxtVille() {
		return txtVille.getText();
	}

	public String getTxtCP() {
		return txtCP.getText();
	}

	public String getTxtMontantAttest() {
		return txtMontantAttest.getText();
	}

	public JComboBox getCmbtitre() {
		return cmbtitre;
	}

	public JDateChooser getDateChooser() {
		return dateChooser;
	}
	
	public void setDateChooser(JDateChooser dateChooser) {
		this.dateChooser = dateChooser;
	}
	/**
	 * Création du Frame 
	 */
	public View() {
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
		
		JComboBox cmbTitre = new JComboBox();
		cmbTitre.setModel(new DefaultComboBoxModel(new String[] {"Madame", "Mademoiselle", "Monsieur", "Aucun titre"}));
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
		
		// TODO voir pour un auto-remplissage du CP une fois la ville renseignée 
		
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
		
		// TODO corriger le bug lié au getDateChooser
		
		JDateChooser dateChooser = new JDateChooser(); 
		dateChooser.setDateFormatString("dd MMMM yyyy");
		//dateChooser.setCalendar(Calendar.getInstance());
		dateChooser.setBounds(240, 310, 153, 20);
		contentPane.add(dateChooser);
		
		// TODO event escape + enter sur quitter et enregistrer respectivement -> enlever l'utilisation de la touche escape dans les deux cas
		// TODO customiser les boutons 
		
		/**
		 * Bouton enregistrer
		 * Event sur le clic et key enter	
		 */
		JButton btnEnregistrer = new JButton("Enregistrer");
		// Méthode isInputValid() lors de l'event key enter
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
		// Méthode isInputValid() lors de l'event clic button enregistrer
		btnEnregistrer.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					isInputValid();
				} catch (InvalidFormatException e1) {
					e1.printStackTrace();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		});
		btnEnregistrer.setFont(new Font("Tahoma", Font.BOLD, 14));
		btnEnregistrer.setBounds(34, 370, 153, 48);
		contentPane.add(btnEnregistrer);
		
		/**
		 * Bouton quitter
		 * Event sur le clic bouton et escape key
		 */
		JButton btnQuitter = new JButton("Quitter");
		// Méthode close() lors de l'event key escape
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
		// Méthode close() lors de l'event clic button quitter
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

// TODO essayer de configurer également les keys messagebox