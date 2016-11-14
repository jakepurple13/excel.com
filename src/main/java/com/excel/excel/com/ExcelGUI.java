package com.excel.excel.com;

import java.awt.Color;
import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.JFileChooser;
import javax.swing.SpringLayout;
import javax.swing.JButton;
import javax.swing.JLabel;

public class ExcelGUI extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ExcelGUI frame = new ExcelGUI();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	String fileName = "";
	String outWriteName = "";
	ExcelMethods em;
	boolean real = false;
	boolean output = false;
	JButton btnBegin;
	JLabel lblOutput;
	JLabel lblNewLabel;

	/**
	 * Create the frame.
	 * @throws IOException 
	 */
	public ExcelGUI() throws IOException {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		SpringLayout sl_contentPane = new SpringLayout();
		contentPane.setLayout(sl_contentPane);

		lblNewLabel = new JLabel("Real Time File");
		contentPane.add(lblNewLabel);

		JButton btnRealTimeFile = new JButton("Real Time File Chooser");
		sl_contentPane.putConstraint(SpringLayout.NORTH, lblNewLabel, 6, SpringLayout.SOUTH, btnRealTimeFile);
		sl_contentPane.putConstraint(SpringLayout.WEST, lblNewLabel, 10, SpringLayout.WEST, btnRealTimeFile);
		sl_contentPane.putConstraint(SpringLayout.NORTH, btnRealTimeFile, 10, SpringLayout.NORTH, contentPane);
		sl_contentPane.putConstraint(SpringLayout.WEST, btnRealTimeFile, 10, SpringLayout.WEST, contentPane);
		contentPane.add(btnRealTimeFile);

		JButton btnOutputFile = new JButton("Output File");
		sl_contentPane.putConstraint(SpringLayout.WEST, btnOutputFile, 0, SpringLayout.WEST, btnRealTimeFile);
		contentPane.add(btnOutputFile);

		btnBegin = new JButton("Begin");
		sl_contentPane.putConstraint(SpringLayout.EAST, lblNewLabel, 0, SpringLayout.EAST, btnBegin);
		sl_contentPane.putConstraint(SpringLayout.SOUTH, btnBegin, -10, SpringLayout.SOUTH, contentPane);
		sl_contentPane.putConstraint(SpringLayout.NORTH, btnOutputFile, 0, SpringLayout.NORTH, btnBegin);
		sl_contentPane.putConstraint(SpringLayout.EAST, btnBegin, -10, SpringLayout.EAST, contentPane);
		contentPane.add(btnBegin);

		lblOutput = new JLabel("Output");
		sl_contentPane.putConstraint(SpringLayout.WEST, lblOutput, 20, SpringLayout.WEST, contentPane);
		sl_contentPane.putConstraint(SpringLayout.SOUTH, lblOutput, -6, SpringLayout.NORTH, btnOutputFile);
		sl_contentPane.putConstraint(SpringLayout.EAST, lblOutput, 0, SpringLayout.EAST, lblNewLabel);
		contentPane.add(lblOutput);

		em = new ExcelMethods();

		btnBegin.setEnabled(false);

		btnRealTimeFile.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser chooser = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel file", "xlsx");
				chooser.setFileFilter(filter);
				chooser.setCurrentDirectory(new File(System.getProperty("user.home"), "Downloads"));
				int returnVal = chooser.showOpenDialog(contentPane);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					System.out.println("You chose to open this file: " + chooser.getSelectedFile().getAbsolutePath());
					fileName = chooser.getSelectedFile().getAbsolutePath();

					lblNewLabel.setText(chooser.getSelectedFile().getAbsolutePath());

					real = true;

				}

				isEverythingGood();

			}
		});

		btnOutputFile.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser chooser = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel", "xlsx", "xls");
				chooser.setFileFilter(filter);
				chooser.addChoosableFileFilter(filter);
				chooser.setAcceptAllFileFilterUsed(true);
				chooser.setDragEnabled(true);
				chooser.setCurrentDirectory(new File(System.getProperty("user.home"), "Desktop"));
				int returnVal = chooser.showSaveDialog(contentPane);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					System.out.println(
							"You chose to open this file: " + chooser.getSelectedFile().getAbsolutePath() + ".xlsx");
					outWriteName = chooser.getSelectedFile().getAbsolutePath() + ".xlsx";

					lblOutput.setText(chooser.getSelectedFile().getAbsolutePath() + ".xlsx");

					output = true;
				}

				isEverythingGood();

			}
		});

		btnBegin.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				em.actualIn(fileName);
				em.outWrite(outWriteName);

				contentPane.setBackground(Color.GREEN);
			}
		});

	}

	public void isEverythingGood() {
		if (real && output) {
			btnBegin.setEnabled(true);
		}
	}

}
