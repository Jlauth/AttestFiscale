import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JLabel;
import javax.swing.JTextField;
import com.toedter.calendar.JDateChooser;
import javax.swing.JButton;
import java.awt.Font;
import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;
import java.awt.event.ActionListener;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Calendar;
import java.util.Date;
import java.awt.event.ActionEvent;

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
	
	/**
	 * méthode fermeture de l'application
	 */
	public void close() {
		dispose();
	}
	
	/**
	 * méthode bouton enregistrer
	 */
	public void save() {
		Attestation attestation = new Attestation(this);
		attestation.saveDoc();
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
		cmbTitre.setModel(new DefaultComboBoxModel(new String[] {"Madame", "Monsieur", "Autre"}));
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
		Calendar c = Calendar.getInstance();
		c.add(Calendar.YEAR, -1);
		JDateChooser chooser = new JDateChooser(c.getTime());
		this.add(chooser);
		dateChooser.setBounds(240, 310, 153, 20);
		contentPane.add(dateChooser);
		
		/**
		 * Bouton enregistrer	
		 */
		JButton btnEnregistrer = new JButton("Enregistrer");
		btnEnregistrer.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				save();
			}
		});
		btnEnregistrer.setFont(new Font("Tahoma", Font.BOLD, 14));
		btnEnregistrer.setBounds(34, 370, 153, 48);
		contentPane.add(btnEnregistrer);
		
		/**
		 * Bouton quitter
		 */
		JButton btnQuitter = new JButton("Quitter");
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
