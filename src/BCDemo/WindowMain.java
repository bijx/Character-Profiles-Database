package BCDemo;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JScrollPane;
import javax.swing.JList;
import javax.swing.JButton;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import java.awt.GridLayout;
import java.awt.TextField;
import java.util.ArrayList;

import javax.swing.JTextPane;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.JMenuBar;
import javax.swing.JMenu;
import java.awt.event.ActionListener;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.awt.event.ActionEvent;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.border.BevelBorder;
import javax.swing.border.CompoundBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.border.LineBorder;
import java.awt.Color;
import javax.swing.border.MatteBorder;
import javax.swing.border.TitledBorder;

import jxl.CellFeatures;
import jxl.CellType;
import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.write.WritableCell;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import javax.swing.JSpinner;
import javax.swing.SpinnerDateModel;
import java.util.Date;
import java.util.Scanner;
import java.util.Calendar;
import java.awt.Label;
import javax.swing.JComboBox;
import java.awt.Button;
import java.awt.Toolkit;

public class WindowMain implements ActionListener, WindowListener {

	static String lastSelected = "";

	static ArrayList<Character> chars = new ArrayList<Character>();

	static JTextPane textPane = new JTextPane();

	private JFrame frmCharacterProfiles;
	private static JTextField nameField;
	private static JTextField ageField;
	private static JTextField raceField;
	private static JList list = new JList();
	private static JTextField eyeColField;
	private static JTextField hairColField;
	private static JTextField movieField;
	private static JTextField favColField;
	private static JTextField favBookField;
	private static JTextField supNatField;
	private static JTextField DOBField;
	private static JTextField momField;
	private static JTextField dadField;
	private static JTextField broField;
	private static JTextField sisField;
	private static JTextField loveField;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					WindowMain window = new WindowMain();
					window.frmCharacterProfiles.setVisible(true);
					try {
						loadFromFile();
					} catch (FileNotFoundException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});

	}

	/**
	 * Create the application.
	 * 
	 * @throws UnsupportedLookAndFeelException
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws ClassNotFoundException
	 */
	public WindowMain() throws ClassNotFoundException, InstantiationException, IllegalAccessException,
			UnsupportedLookAndFeelException {

		UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmCharacterProfiles = new JFrame();
		frmCharacterProfiles.setIconImage(Toolkit.getDefaultToolkit().getImage(WindowMain.class.getResource("/BCDemo/Address-Book-icon.png")));
		frmCharacterProfiles.addWindowListener(this);
		frmCharacterProfiles.setResizable(false);
		frmCharacterProfiles.setTitle("Character Profiles");
		frmCharacterProfiles.setBounds(100, 100, 648, 555);
		frmCharacterProfiles.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmCharacterProfiles.getContentPane().setLayout(null);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(10, 11, 142, 439);
		frmCharacterProfiles.getContentPane().add(scrollPane);
		list.setToolTipText("List of all entered characters.");

		scrollPane.setViewportView(list);

		JButton btnOpen = new JButton("Open");
		btnOpen.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				loadCharacter();
			}
		});
		btnOpen.setBounds(10, 461, 142, 23);
		frmCharacterProfiles.getContentPane().add(btnOpen);

		JPanel panel = new JPanel();
		panel.setBounds(162, 11, 460, 439);
		frmCharacterProfiles.getContentPane().add(panel);
		panel.setLayout(new GridLayout(0, 1, 0, 0));

		JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);
		panel.add(tabbedPane);

		JPanel panel_1 = new JPanel();
		tabbedPane.addTab("Information", null, panel_1, "General information on the character, and base statistics.");
		panel_1.setLayout(null);

		JPanel panel_3 = new JPanel();
		panel_3.setToolTipText("Basic info about the character, usually consists of data found on a piece of ID.");
		panel_3.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Basic Info",
				TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_3.setBounds(10, 11, 435, 90);
		panel_1.add(panel_3);
		panel_3.setLayout(null);

		nameField = new JTextField();
		nameField.setBounds(10, 41, 145, 20);
		panel_3.add(nameField);
		nameField.setColumns(10);

		JLabel lblName = new JLabel("Name");
		lblName.setBounds(10, 23, 46, 14);
		panel_3.add(lblName);

		JLabel lblAge = new JLabel("Age");
		lblAge.setBounds(165, 23, 46, 14);
		panel_3.add(lblAge);

		ageField = new JTextField();
		ageField.setBounds(165, 41, 51, 20);
		panel_3.add(ageField);
		ageField.setColumns(10);

		JLabel lblDateOfBirth = new JLabel("Date Of Birth");
		lblDateOfBirth.setBounds(237, 23, 136, 14);
		panel_3.add(lblDateOfBirth);

		DOBField = new JTextField();
		DOBField.setColumns(10);
		DOBField.setBounds(233, 41, 192, 20);
		panel_3.add(DOBField);

		JPanel panel_4 = new JPanel();
		panel_4.setLayout(null);
		panel_4.setToolTipText("Physical appearance.");
		panel_4.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Visual Traits",
				TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_4.setBounds(10, 112, 435, 88);
		panel_1.add(panel_4);

		eyeColField = new JTextField();
		eyeColField.setColumns(10);
		eyeColField.setBounds(6, 30, 119, 20);
		panel_4.add(eyeColField);

		JLabel lblEyeColour = new JLabel("Eye Colour");
		lblEyeColour.setBounds(10, 16, 115, 14);
		panel_4.add(lblEyeColour);

		JLabel lblHairColour = new JLabel("Hair Colour");
		lblHairColour.setBounds(135, 12, 82, 14);
		panel_4.add(lblHairColour);

		hairColField = new JTextField();
		hairColField.setColumns(10);
		hairColField.setBounds(135, 30, 117, 20);
		panel_4.add(hairColField);

		raceField = new JTextField();
		raceField.setBounds(255, 30, 170, 20);
		panel_4.add(raceField);
		raceField.setColumns(10);

		JLabel lblBehaviour = new JLabel("Race");
		lblBehaviour.setBounds(257, 12, 84, 14);
		panel_4.add(lblBehaviour);

		JPanel panel_5 = new JPanel();
		panel_5.setLayout(null);
		panel_5.setToolTipText("Other useful information.");
		panel_5.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Other Information",
				TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_5.setBounds(10, 213, 435, 124);
		panel_1.add(panel_5);

		movieField = new JTextField();
		movieField.setColumns(10);
		movieField.setBounds(6, 30, 119, 20);
		panel_5.add(movieField);

		JLabel lblFavouriteMovie = new JLabel("Favourite Movie");
		lblFavouriteMovie.setBounds(10, 16, 115, 14);
		panel_5.add(lblFavouriteMovie);

		JLabel lblFavouriteColour = new JLabel("Favourite Colour");
		lblFavouriteColour.setBounds(135, 16, 119, 14);
		panel_5.add(lblFavouriteColour);

		favColField = new JTextField();
		favColField.setColumns(10);
		favColField.setBounds(135, 30, 117, 20);
		panel_5.add(favColField);

		favBookField = new JTextField();
		favBookField.setColumns(10);
		favBookField.setBounds(255, 30, 170, 20);
		panel_5.add(favBookField);

		JLabel lblFavouriteBook = new JLabel("Favourite Book");
		lblFavouriteBook.setBounds(257, 16, 155, 14);
		panel_5.add(lblFavouriteBook);

		JLabel lblSupernaturalAbility = new JLabel("Supernatural Ability");
		lblSupernaturalAbility.setBounds(47, 83, 119, 14);
		panel_5.add(lblSupernaturalAbility);

		supNatField = new JTextField();
		supNatField.setColumns(10);
		supNatField.setBounds(163, 80, 185, 20);
		panel_5.add(supNatField);

		JPanel panel_2 = new JPanel();
		tabbedPane.addTab("Bio", null, panel_2, "The biography of the selected character.");
		panel_2.setLayout(new GridLayout(0, 1, 0, 0));

		JScrollPane scrollPane_1 = new JScrollPane();
		panel_2.add(scrollPane_1);

		scrollPane_1.setViewportView(textPane);

		JPanel panel_6 = new JPanel();
		tabbedPane.addTab("Relationships", null, panel_6,
				"Relationships with other people (can be characters in the profiles editor)");
		panel_6.setLayout(null);

		Label label = new Label("Mother:");
		label.setBounds(10, 10, 62, 22);
		panel_6.add(label);

		Label label_1 = new Label("Father:");
		label_1.setBounds(10, 38, 62, 22);
		panel_6.add(label_1);

		Label label_2 = new Label("Brother:");
		label_2.setBounds(10, 66, 62, 22);
		panel_6.add(label_2);

		Label label_3 = new Label("Sister:");
		label_3.setBounds(10, 94, 62, 22);
		panel_6.add(label_3);

		Label label_4 = new Label("Love Interest:");
		label_4.setBounds(10, 122, 87, 22);
		panel_6.add(label_4);

		Button button = new Button("Goto");
		button.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				search(momField.getText());
			}
		});
		button.setBounds(324, 10, 50, 20);
		panel_6.add(button);

		momField = new JTextField();
		momField.setBounds(119, 10, 199, 20);
		panel_6.add(momField);
		momField.setColumns(10);

		dadField = new JTextField();
		dadField.setColumns(10);
		dadField.setBounds(119, 38, 199, 20);
		panel_6.add(dadField);

		broField = new JTextField();
		broField.setColumns(10);
		broField.setBounds(119, 68, 199, 20);
		panel_6.add(broField);

		sisField = new JTextField();
		sisField.setColumns(10);
		sisField.setBounds(119, 94, 199, 20);
		panel_6.add(sisField);

		loveField = new JTextField();
		loveField.setColumns(10);
		loveField.setBounds(119, 124, 199, 20);
		panel_6.add(loveField);

		Button button_1 = new Button("Goto");
		button_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				search(dadField.getText());
			}
		});
		button_1.setBounds(324, 38, 50, 20);
		panel_6.add(button_1);

		Button button_2 = new Button("Goto");
		button_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				search(broField.getText());
			}
		});
		button_2.setBounds(324, 68, 50, 20);
		panel_6.add(button_2);

		Button button_3 = new Button("Goto");
		button_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				search(sisField.getText());
			}
		});
		button_3.setBounds(324, 96, 50, 20);
		panel_6.add(button_3);

		Button button_4 = new Button("Goto");
		button_4.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				search(loveField.getText());
			}
		});
		button_4.setBounds(324, 124, 50, 20);
		panel_6.add(button_4);

		JButton btnSave = new JButton("Save");
		btnSave.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (!nameField.getText().equals("")) {
					save();
				} else {
					JOptionPane.showMessageDialog(null, "Please enter a name.", "Save Error",
							JOptionPane.ERROR_MESSAGE);
				}

			}
		});
		btnSave.setBounds(162, 461, 128, 23);
		frmCharacterProfiles.getContentPane().add(btnSave);

		JButton btnUpdate = new JButton("Update");
		btnUpdate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (!nameField.getText().equals("")) {

					for (int i = 0; i < chars.size(); i++) {
						if (chars.get(i).getName().equals(lastSelected)) {
							chars.get(i).updateCharacter(nameField.getText(), ageField.getText(), DOBField.getText(),
									textPane.getText(), eyeColField.getText(), hairColField.getText(),
									raceField.getText(), movieField.getText(), favColField.getText(),
									favBookField.getText(), supNatField.getText(), momField.getText(),
									dadField.getText(), broField.getText(), sisField.getText(), loveField.getText());
							JOptionPane.showMessageDialog(null,
									"Successfully updated " + chars.get(i).getName() + "'s information!");
						}
					}
					updateList();
					clearFields();

				} else {
					JOptionPane.showMessageDialog(null, "Please enter a name.", "Save Error",
							JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		btnUpdate.setBounds(300, 461, 128, 23);
		frmCharacterProfiles.getContentPane().add(btnUpdate);

		JButton btnDelete = new JButton("Delete");
		btnDelete.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				delete();
			}
		});
		btnDelete.setBounds(438, 461, 128, 23);
		frmCharacterProfiles.getContentPane().add(btnDelete);

		JMenuBar menuBar = new JMenuBar();
		frmCharacterProfiles.setJMenuBar(menuBar);

		JMenu mnFile = new JMenu("File");
		menuBar.add(mnFile);

		JMenuItem mntmSave = new JMenuItem("Save");
		mnFile.add(mntmSave);

		JMenu mnExport = new JMenu("Export");
		mnFile.add(mnExport);

		JMenuItem mntmExcel = new JMenuItem("Excel");
		mnExport.add(mntmExcel);

		JMenu mnTools = new JMenu("Tools");
		menuBar.add(mnTools);

		JMenuItem mntmClearFields = new JMenuItem("Clear fields");
		mnTools.add(mntmClearFields);
		
		JMenu mnAbout = new JMenu("Help");
		menuBar.add(mnAbout);
		
		JMenuItem mntmAbout = new JMenuItem("About");
		mnAbout.add(mntmAbout);

		mntmAbout.addActionListener(this);
		mntmExcel.addActionListener(this);
		mntmClearFields.addActionListener(this);
		mntmSave.addActionListener(this);
	}

	public static void search(String n) {
		for (int i = 0; i < chars.size(); i++) {
			if (chars.get(i).getName().equals(n)) {
				nameField.setText(chars.get(i).getName());
				ageField.setText(chars.get(i).getAge());
				DOBField.setText(chars.get(i).getBehav());
				textPane.setText(chars.get(i).getBio());
				eyeColField.setText(chars.get(i).getEyeCol());
				hairColField.setText(chars.get(i).getHairCol());
				raceField.setText(chars.get(i).getRace());
				movieField.setText(chars.get(i).getFavMovie());
				favColField.setText(chars.get(i).getFavCol());
				favBookField.setText(chars.get(i).getFavBook());
				supNatField.setText(chars.get(i).getSuper());
				momField.setText(chars.get(i).getMom());
				dadField.setText(chars.get(i).getDad());
				broField.setText(chars.get(i).getBro());
				sisField.setText(chars.get(i).getSis());
				loveField.setText(chars.get(i).getLove());

				lastSelected = chars.get(i).getName();

				return;
			}
		}

		JOptionPane.showMessageDialog(null, "Could not load character.", "Load Error", JOptionPane.ERROR_MESSAGE);

	}

	public static void delete() {
		if (lastSelected.equals("")) {
			JOptionPane.showMessageDialog(null, "Please load a user to delete.", "Delete Error",
					JOptionPane.ERROR_MESSAGE);
		} else {
			int n = JOptionPane.showConfirmDialog(null, "Are you sure you wish to delete " + lastSelected + "?",
					"Confirm Delete", JOptionPane.YES_NO_OPTION);
			if (n == JOptionPane.YES_OPTION) {
				for (int i = 0; i < chars.size(); i++) {
					if (chars.get(i).getName().equals(lastSelected)) {
						chars.remove(i);
						updateList();
						JOptionPane.showMessageDialog(null, "User deleted!", "Deleted!", JOptionPane.ERROR_MESSAGE);
						lastSelected = "";
						clearFields();
						return;
					}
				}
			}
		}

	}

	public static void save() {

		Character newChar = new Character(nameField.getText(), ageField.getText(), DOBField.getText(),
				textPane.getText(), eyeColField.getText(), hairColField.getText(), raceField.getText(),
				movieField.getText(), favColField.getText(), favBookField.getText(), supNatField.getText(),
				momField.getText(), dadField.getText(), broField.getText(), sisField.getText(), loveField.getText());
		chars.add(newChar);
		updateList();
		JOptionPane.showMessageDialog(null, "Successfully saved " + nameField.getText() + "'s information!");
		clearFields();
	}

	public static void clearFields() {
		nameField.setText("");
		ageField.setText("");
		DOBField.setText("");
		textPane.setText("");
		eyeColField.setText("");
		hairColField.setText("");
		raceField.setText("");
		movieField.setText("");
		favColField.setText("");
		favBookField.setText("");
		supNatField.setText("");
		momField.setText("");
		dadField.setText("");
		broField.setText("");
		sisField.setText("");
		loveField.setText("");
	}

	public static void updateList() {
		String[] charArray = new String[chars.size()];
		for (int i = 0; i < chars.size(); i++) {
			charArray[i] = chars.get(i).getName();
		}

		sort(charArray);

		list.setListData(charArray);
		list.updateUI();
	}

	public static void sort(String[] arr) {
		String temp;
		for (int i = 0; i < arr.length - 1; i++) {
			for (int j = 0; j < arr.length - 1; j++) {
				if (arr[j].compareToIgnoreCase(arr[j + 1]) > 0) {
					temp = arr[j];
					arr[j] = arr[j + 1];
					arr[j + 1] = temp;

				}

			}
		}
	}

	public static void loadCharacter() {
		String n = (String) list.getSelectedValue();
		for (int i = 0; i < chars.size(); i++) {
			if (chars.get(i).getName().equals(n)) {
				nameField.setText(chars.get(i).getName());
				ageField.setText(chars.get(i).getAge());
				DOBField.setText(chars.get(i).getBehav());
				textPane.setText(chars.get(i).getBio());
				eyeColField.setText(chars.get(i).getEyeCol());
				hairColField.setText(chars.get(i).getHairCol());
				raceField.setText(chars.get(i).getRace());
				movieField.setText(chars.get(i).getFavMovie());
				favColField.setText(chars.get(i).getFavCol());
				favBookField.setText(chars.get(i).getFavBook());
				supNatField.setText(chars.get(i).getSuper());
				momField.setText(chars.get(i).getMom());
				dadField.setText(chars.get(i).getDad());
				broField.setText(chars.get(i).getBro());
				sisField.setText(chars.get(i).getSis());
				loveField.setText(chars.get(i).getLove());

				lastSelected = chars.get(i).getName();
			}
		}
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		String cmd = e.getActionCommand();

		if (cmd.equals("Save")) {
			try {
				saveToFile();
			} catch (FileNotFoundException e1) {
				JOptionPane.showMessageDialog(null, "Error! Could not save.", "Save Error", JOptionPane.ERROR_MESSAGE);
				e1.printStackTrace();
			}
		} else if (cmd.equals("Clear fields")) {
			clearFields();
		} else if (cmd.equals("Excel")) {
			try {
				exportXLS();
			} catch (RowsExceededException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (WriteException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}else if(cmd.equals("About")){
			JOptionPane.showMessageDialog(null, "Character Profile Database V2.5.0\nLicensed to Erin Chace\n\nCPDv2.5 was made by BijxFilms. All Rights Reserved.", "About", JOptionPane.INFORMATION_MESSAGE);
		}

	}

	public static void saveToFile() throws FileNotFoundException {

		File myFile = new File("characters.cp");
		PrintWriter output = new PrintWriter(myFile);

		for (int i = 0; i < chars.size(); i++) {
			output.println(chars.get(i).getName());
			output.println(chars.get(i).getAge());
			output.println(chars.get(i).getBehav());
			output.println(chars.get(i).getBio());
			output.println(chars.get(i).getEyeCol());
			output.println(chars.get(i).getHairCol());
			output.println(chars.get(i).getRace());
			output.println(chars.get(i).getFavMovie());
			output.println(chars.get(i).getFavCol());
			output.println(chars.get(i).getFavBook());
			output.println(chars.get(i).getSuper());
			output.println(chars.get(i).getMom());
			output.println(chars.get(i).getDad());
			output.println(chars.get(i).getBro());
			output.println(chars.get(i).getSis());
			output.println(chars.get(i).getLove());
		}

		output.close();
	}

	public static void loadFromFile() throws FileNotFoundException {
		File file = new java.io.File("characters.cp");
		Scanner input = new Scanner(file);
		while (input.hasNextLine()) {
			Character newChar = new Character(input.nextLine(), input.nextLine(), input.nextLine(), input.nextLine(),
					input.nextLine(), input.nextLine(), input.nextLine(), input.nextLine(), input.nextLine(),
					input.nextLine(), input.nextLine(), input.nextLine(), input.nextLine(), input.nextLine(),
					input.nextLine(), input.nextLine());
			chars.add(newChar);

		}
		updateList();
		JOptionPane.showMessageDialog(null, "Successfully loaded data!");
		clearFields();
	}

	@Override
	public void windowActivated(WindowEvent e) {
		// TODO Auto-generated method stub

	}

	@Override
	public void windowClosed(WindowEvent e) {
		// TODO Auto-generated method stub

	}

	@Override
	public void windowClosing(WindowEvent e) {
		// TODO Auto-generated method stub
		try {
			saveToFile();
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
	}

	@Override
	public void windowDeactivated(WindowEvent e) {
		// TODO Auto-generated method stub

	}

	@Override
	public void windowDeiconified(WindowEvent e) {
		// TODO Auto-generated method stub

	}

	@Override
	public void windowIconified(WindowEvent e) {
		// TODO Auto-generated method stub

	}

	@Override
	public void windowOpened(WindowEvent e) {
		// TODO Auto-generated method stub

	}

	public static void exportXLS() throws RowsExceededException, WriteException, IOException {
		File f = new File("Characters.xls");
		WritableWorkbook book = Workbook.createWorkbook(f);
		WritableSheet sheet = book.createSheet("Characters", 0);

		// Generate Basic Table Headers
		jxl.write.Label l = new jxl.write.Label(0, 0, "Name");
		sheet.addCell(l);
		l = new jxl.write.Label(1, 0, "Age");
		sheet.addCell(l);
		l = new jxl.write.Label(2, 0, "DOB");
		sheet.addCell(l);
		l = new jxl.write.Label(3, 0, "Bio");
		sheet.addCell(l);
		l = new jxl.write.Label(4, 0, "Eye Colour");
		sheet.addCell(l);
		l = new jxl.write.Label(5, 0, "Hair Colour");
		sheet.addCell(l);
		l = new jxl.write.Label(6, 0, "Race");
		sheet.addCell(l);
		l = new jxl.write.Label(7, 0, "Favourite Movie");
		sheet.addCell(l);
		l = new jxl.write.Label(8, 0, "Favourite Colour");
		sheet.addCell(l);
		l = new jxl.write.Label(9, 0, "Favourite Book");
		sheet.addCell(l);
		l = new jxl.write.Label(10, 0, "Supernatural Ability");
		sheet.addCell(l);
		l = new jxl.write.Label(11, 0, "Mother");
		sheet.addCell(l);
		l = new jxl.write.Label(12, 0, "Father");
		sheet.addCell(l);
		l = new jxl.write.Label(13, 0, "Brother");
		sheet.addCell(l);
		l = new jxl.write.Label(14, 0, "Sister");
		sheet.addCell(l);
		l = new jxl.write.Label(15, 0, "Love Interest");
		sheet.addCell(l);
		

		for (int i = 1; i < chars.size()+1; i++) {

			l = new jxl.write.Label(0, i, chars.get(i-1).getName());
			sheet.addCell(l);
			l = new jxl.write.Label(1, i, chars.get(i-1).getAge());
			sheet.addCell(l);
			l = new jxl.write.Label(2, i, chars.get(i-1).getBehav());
			sheet.addCell(l);
			l = new jxl.write.Label(3, i, chars.get(i-1).getBio());
			sheet.addCell(l);
			l = new jxl.write.Label(4, i, chars.get(i-1).getEyeCol());
			sheet.addCell(l);
			l = new jxl.write.Label(5, i, chars.get(i-1).getHairCol());
			sheet.addCell(l);
			l = new jxl.write.Label(6, i, chars.get(i-1).getRace());
			sheet.addCell(l);
			l = new jxl.write.Label(7, i, chars.get(i-1).getFavMovie());
			sheet.addCell(l);
			l = new jxl.write.Label(8, i, chars.get(i-1).getFavCol());
			sheet.addCell(l);
			l = new jxl.write.Label(9, i, chars.get(i-1).getFavBook());
			sheet.addCell(l);
			l = new jxl.write.Label(10, i, chars.get(i-1).getSuper());
			sheet.addCell(l);
			l = new jxl.write.Label(11, i, chars.get(i-1).getMom());
			sheet.addCell(l);
			l = new jxl.write.Label(12, i, chars.get(i-1).getDad());
			sheet.addCell(l);
			l = new jxl.write.Label(13, i, chars.get(i-1).getBro());
			sheet.addCell(l);
			l = new jxl.write.Label(14, i, chars.get(i-1).getSis());
			sheet.addCell(l);
			l = new jxl.write.Label(15, i, chars.get(i-1).getLove());
			sheet.addCell(l);
		}
		
		book.write();
		book.close();

	}
}
