/*
 * Created on 15.06.2005
 *
 * Montpellier RIO Imaging
 * 
 * Homepage: http://www.mri.cnrs.fr/
 * 
 * created by
 * 		Volker Baecker and Pierre Travo
 * 
 * contact:
 * 	volker.baecker@mri.cnrs.fr
 *  	pierre.travo@mri.cnrs.fr
 */

package helloImaris;

import javax.swing.JFrame;

import javax.swing.JPanel;
import javax.swing.JButton;
import javax.swing.JTextField;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.SafeArray;
import com.jacob.com.Variant;
import java.awt.Checkbox;
import java.util.Random;

import javax.swing.JComboBox;
import javax.swing.JLabel;
/**
 * @author Volker
 * 
 * TODO To change the template for this generated type comment go to Window -
 * Preferences - Java - Code Style - Code Templates
 */
public class ImarisRemoteControl extends JFrame {

	private JPanel jPanel = null;

	private JPanel jPanel1 = null;

	private JButton startImarisButton = null;

	private JTextField imagePath1TextField = null;

	private JTextField imagePath2TextField = null;

	private JButton loadImage1Button = null;

	private JButton loadImage2Button = null;

	private JButton exitImarisButton = null;

	private Checkbox checkbox = null;

	private JPanel jPanel2 = null;

	private JComboBox typeComboBox = null;

	private JTextField xTextField = null;

	private JTextField yTextField = null;

	private JTextField zTextField = null;

	private JTextField chTextField = null;

	private JTextField tTextField = null;

	private JLabel jLabel = null;

	private JLabel jLabel1 = null;

	private JLabel jLabel2 = null;

	private JLabel jLabel3 = null;

	private JLabel jLabel4 = null;

	private JLabel jLabel5 = null;

	private JButton getSizeButton = null;

	private JButton setSizeButton = null;

	private JPanel jPanel3 = null;

	private JButton setRandomVoxelIntensitiesButton = null;

	private ActiveXComponent imarisApplication;

	/**
	 * This is the default constructor
	 */
	public ImarisRemoteControl() {
		super();
		initialize();
	}

	/**
	 * This method initializes this
	 * 
	 * @return void
	 */
	private void initialize() {
		this
				.setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
		this.setContentPane(getJPanel());
		this.setSize(440, 545);
		this.setTitle("Hello Imaris DataSet");
		this.addWindowListener(new java.awt.event.WindowAdapter() {
			public void windowClosing(java.awt.event.WindowEvent e) {
				if (imarisApplication != null) {
					imarisApplication.setProperty("mUserControl", !checkbox
							.getState());
				}
			}
		});
	}

	/**
	 * This method initializes jPanel
	 * 
	 * @return javax.swing.JPanel
	 */
	private JPanel getJPanel() {
		if (jPanel == null) {
			jPanel = new JPanel();
			jPanel.setLayout(null);
			jPanel.add(getJPanel1(), null);
			jPanel.add(getJPanel2(), null);
			jPanel.add(getJPanel3(), null);
		}
		return jPanel;
	}

	/**
	 * This method initializes jPanel1
	 * 
	 * @return javax.swing.JPanel
	 */
	private JPanel getJPanel1() {
		if (jPanel1 == null) {
			jPanel1 = new JPanel();
			jPanel1.setLayout(null);
			jPanel1.setBounds(11, 8, 400, 194);
			jPanel1.setBorder(javax.swing.BorderFactory.createTitledBorder(
					null, "Application",
					javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION,
					javax.swing.border.TitledBorder.DEFAULT_POSITION, null,
					null));
			jPanel1.add(getStartImarisButton(), null);
			jPanel1.add(getImagePath1TextField(), null);
			jPanel1.add(getImagePath2TextField(), null);
			jPanel1.add(getExitImarisButton(), null);
			jPanel1.add(getCheckbox(), null);
			jPanel1.add(getLoadImage1Button(), null);
			jPanel1.add(getLoadImage2Button(), null);
		}
		return jPanel1;
	}

	/**
	 * This method initializes startImarisButton
	 * 
	 * @return javax.swing.JButton
	 */
	private JButton getStartImarisButton() {
		if (startImarisButton == null) {
			startImarisButton = new JButton();
			startImarisButton.setBounds(16, 22, 364, 20);
			startImarisButton.setText("Start Imaris");
			startImarisButton.setBorder(javax.swing.BorderFactory
					.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
			startImarisButton
					.addActionListener(new java.awt.event.ActionListener() {
						public void actionPerformed(java.awt.event.ActionEvent e) {
							if (imarisApplication == null) {
								imarisApplication = new ActiveXComponent(
										"Imaris.Application");
                imarisApplication.setProperty("mVisible", true);
							}
						}
					});
		}
		return startImarisButton;
	}

	/**
	 * This method initializes imagePath1TextField
	 * 
	 * @return javax.swing.JTextField
	 */
	private JTextField getImagePath1TextField() {
		if (imagePath1TextField == null) {
			imagePath1TextField = new JTextField();
			imagePath1TextField.setBounds(16, 62, 276, 20);
			imagePath1TextField
					.setText("c:\\program files\\bitplane\\imaris 4\\images\\retina.ims");
		}
		return imagePath1TextField;
	}

	/**
	 * This method initializes imagePath2TextField
	 * 
	 * @return javax.swing.JTextField
	 */
	private JTextField getImagePath2TextField() {
		if (imagePath2TextField == null) {
			imagePath2TextField = new JTextField();
			imagePath2TextField.setBounds(16, 87, 276, 20);
			imagePath2TextField
					.setText("c:\\program files\\bitplane\\imaris 4\\images\\R18Demo.ims");
		}
		return imagePath2TextField;
	}

	/**
	 * This method initializes loadImage1Button
	 * 
	 * @return javax.swing.JButton
	 */
	private JButton getLoadImage1Button() {
		if (loadImage1Button == null) {
			loadImage1Button = new JButton();
			loadImage1Button.setBounds(305, 62, 76, 20);
			loadImage1Button.setText("Load");
			loadImage1Button.setBorder(javax.swing.BorderFactory
					.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
			loadImage1Button
					.addActionListener(new java.awt.event.ActionListener() {
						public void actionPerformed(java.awt.event.ActionEvent e) {
							if (imarisApplication != null) {
								Variant parameter1 = new Variant(
										imagePath1TextField.getText());
								Variant parameter2 = new Variant("");
								imarisApplication.invoke("FileOpen",
										parameter1, parameter2);
							}
						}
					});
		}
		return loadImage1Button;
	}

	/**
	 * This method initializes loadImage2Button
	 * 
	 * @return javax.swing.JButton
	 */
	private JButton getLoadImage2Button() {
		if (loadImage2Button == null) {
			loadImage2Button = new JButton();
			loadImage2Button.setBounds(305, 87, 76, 20);
			loadImage2Button.setText("Load");
			loadImage2Button.setBorder(javax.swing.BorderFactory
					.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
			loadImage2Button
					.addActionListener(new java.awt.event.ActionListener() {
						public void actionPerformed(java.awt.event.ActionEvent e) {
							if (imarisApplication != null) {
								Variant parameter1 = new Variant(
										imagePath2TextField.getText());
								Variant parameter2 = new Variant("");
								imarisApplication.invoke("FileOpen",
										parameter1, parameter2);
							}
						}
					});
		}
		return loadImage2Button;
	}

	/**
	 * This method initializes exitImarisButton
	 * 
	 * @return javax.swing.JButton
	 */
	private JButton getExitImarisButton() {
		if (exitImarisButton == null) {
			exitImarisButton = new JButton();
			exitImarisButton.setBounds(17, 124, 364, 20);
			exitImarisButton.setBorder(javax.swing.BorderFactory
					.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
			exitImarisButton.setText("Exit Imaris");
			exitImarisButton
					.addActionListener(new java.awt.event.ActionListener() {
						public void actionPerformed(java.awt.event.ActionEvent e) {
							if (imarisApplication != null) {
								imarisApplication.safeRelease();
								imarisApplication = null;
							}
						}
					});
		}
		return exitImarisButton;
	}

	/**
	 * This method initializes checkbox
	 * 
	 * @return java.awt.Checkbox
	 */
	private Checkbox getCheckbox() {
		if (checkbox == null) {
			checkbox = new Checkbox();
			checkbox.setLabel("Exit Imaris when closing this Dialog");
			checkbox.setBounds(17, 154, 243, 23);
			checkbox.setState(true);
			checkbox.addItemListener(new java.awt.event.ItemListener() {
				public void itemStateChanged(java.awt.event.ItemEvent e) {
					if (imarisApplication != null) {
						imarisApplication.setProperty("mUserControl", !checkbox
								.getState());
					}
				}
			});
		}
		return checkbox;
	}

	/**
	 * This method initializes jPanel2
	 * 
	 * @return javax.swing.JPanel
	 */
	private JPanel getJPanel2() {
		if (jPanel2 == null) {
			jPanel2 = new JPanel();
			jLabel = new JLabel();
			jLabel1 = new JLabel();
			jLabel2 = new JLabel();
			jLabel3 = new JLabel();
			jLabel4 = new JLabel();
			jLabel5 = new JLabel();
			jPanel2.setLayout(null);
			jPanel2.setBounds(11, 210, 400, 226);
			jPanel2.setBorder(javax.swing.BorderFactory.createTitledBorder(
					null, "DataSet Size",
					javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION,
					javax.swing.border.TitledBorder.DEFAULT_POSITION, null,
					null));
			jLabel.setBounds(82, 15, 44, 16);
			jLabel.setText("Type");
			jLabel.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
			jLabel1.setBounds(82, 50, 44, 16);
			jLabel1.setText("X");
			jLabel1.setHorizontalTextPosition(javax.swing.SwingConstants.RIGHT);
			jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.TRAILING);
			jLabel2.setBounds(82, 85, 44, 16);
			jLabel2.setText("Y");
			jLabel2.setHorizontalAlignment(javax.swing.SwingConstants.TRAILING);
			jLabel2.setHorizontalTextPosition(javax.swing.SwingConstants.RIGHT);
			jLabel3.setBounds(82, 120, 44, 16);
			jLabel3.setText("Z");
			jLabel3.setHorizontalAlignment(javax.swing.SwingConstants.TRAILING);
			jLabel3.setHorizontalTextPosition(javax.swing.SwingConstants.RIGHT);
			jLabel4.setBounds(82, 155, 44, 16);
			jLabel4.setText("Ch");
			jLabel4.setHorizontalAlignment(javax.swing.SwingConstants.TRAILING);
			jLabel4.setHorizontalTextPosition(javax.swing.SwingConstants.RIGHT);
			jLabel5.setBounds(82, 190, 44, 16);
			jLabel5.setText("T");
			jLabel5.setHorizontalAlignment(javax.swing.SwingConstants.TRAILING);
			jLabel5.setHorizontalTextPosition(javax.swing.SwingConstants.RIGHT);
			jPanel2.add(getTypeComboBox(), null);
			jPanel2.add(getXTextField(), null);
			jPanel2.add(getYTextField(), null);
			jPanel2.add(getZTextField(), null);
			jPanel2.add(getChTextField(), null);
			jPanel2.add(getTTextField(), null);
			jPanel2.add(jLabel, null);
			jPanel2.add(jLabel1, null);
			jPanel2.add(jLabel2, null);
			jPanel2.add(jLabel3, null);
			jPanel2.add(jLabel4, null);
			jPanel2.add(jLabel5, null);
			jPanel2.add(getGetSizeButton(), null);
			jPanel2.add(getSetSizeButton(), null);
		}
		return jPanel2;
	}

	/**
	 * This method initializes typeComboBox
	 * 
	 * @return javax.swing.JComboBox
	 */
	private JComboBox getTypeComboBox() {
		if (typeComboBox == null) {
			String[] choices = { "Unknown", "UInt8", "UInt16", "Float" };
			typeComboBox = new JComboBox(choices);
			typeComboBox.setBounds(132, 12, 121, 23);
			typeComboBox.setEditable(false);
			typeComboBox.setBorder(javax.swing.BorderFactory
					.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
		}
		return typeComboBox;
	}

	/**
	 * This method initializes xTextField
	 * 
	 * @return javax.swing.JTextField
	 */
	private JTextField getXTextField() {
		if (xTextField == null) {
			xTextField = new JTextField();
			xTextField.setBounds(132, 47, 121, 23);
		}
		return xTextField;
	}

	/**
	 * This method initializes yTextField
	 * 
	 * @return javax.swing.JTextField
	 */
	private JTextField getYTextField() {
		if (yTextField == null) {
			yTextField = new JTextField();
			yTextField.setBounds(132, 82, 121, 23);
		}
		return yTextField;
	}

	/**
	 * This method initializes zTextField
	 * 
	 * @return javax.swing.JTextField
	 */
	private JTextField getZTextField() {
		if (zTextField == null) {
			zTextField = new JTextField();
			zTextField.setBounds(132, 117, 121, 23);
		}
		return zTextField;
	}

	/**
	 * This method initializes chTextField
	 * 
	 * @return javax.swing.JTextField
	 */
	private JTextField getChTextField() {
		if (chTextField == null) {
			chTextField = new JTextField();
			chTextField.setBounds(132, 152, 121, 23);
		}
		return chTextField;
	}

	/**
	 * This method initializes tTextField
	 * 
	 * @return javax.swing.JTextField
	 */
	private JTextField getTTextField() {
		if (tTextField == null) {
			tTextField = new JTextField();
			tTextField.setBounds(132, 187, 121, 23);
		}
		return tTextField;
	}

	/**
	 * This method initializes getSizeButton
	 * 
	 * @return javax.swing.JButton
	 */
	private JButton getGetSizeButton() {
		if (getSizeButton == null) {
			getSizeButton = new JButton();
			getSizeButton.setBounds(299, 13, 93, 20);
			getSizeButton.setText("Get Size");
			getSizeButton.setBorder(javax.swing.BorderFactory
					.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
			getSizeButton
					.addActionListener(new java.awt.event.ActionListener() {
						public void actionPerformed(java.awt.event.ActionEvent e) {
							ActiveXComponent dataSet = imarisApplication
									.getPropertyAsComponent("mDataSet");
							int type = dataSet.getPropertyAsInt("mType");
							typeComboBox.setSelectedIndex(type);
							xTextField.setText(dataSet
									.getPropertyAsString("mSizeX"));
							yTextField.setText(dataSet
									.getPropertyAsString("mSizeY"));
							zTextField.setText(dataSet
									.getPropertyAsString("mSizeZ"));
							chTextField.setText(dataSet
									.getPropertyAsString("mSizeC"));
							tTextField.setText(dataSet
									.getPropertyAsString("mSizeT"));
						}
					});
		}
		return getSizeButton;
	}

	/**
	 * This method initializes setSizeButton
	 * 
	 * @return javax.swing.JButton
	 */
	private JButton getSetSizeButton() {
		if (setSizeButton == null) {
			setSizeButton = new JButton();
			setSizeButton.setBounds(299, 48, 93, 20);
			setSizeButton.setText("Set Size");
			setSizeButton.setBorder(javax.swing.BorderFactory
					.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
			setSizeButton
					.addActionListener(new java.awt.event.ActionListener() {
						public void actionPerformed(java.awt.event.ActionEvent e) {
							ActiveXComponent dataSet = imarisApplication
									.getPropertyAsComponent("mDataSet");
							Variant[] parameter = new Variant[6];
							parameter[0] = new Variant(typeComboBox
									.getSelectedIndex());
							parameter[1] = new Variant(Integer.valueOf(
									xTextField.getText()).intValue());
							parameter[2] = new Variant(Integer.valueOf(
									yTextField.getText()).intValue());
							parameter[3] = new Variant(Integer.valueOf(
									zTextField.getText()).intValue());
							parameter[4] = new Variant(Integer.valueOf(
									chTextField.getText()).intValue());
							parameter[5] = new Variant(Integer.valueOf(
									tTextField.getText()).intValue());
							dataSet.invoke("Create", parameter);
						}
					});
		}
		return setSizeButton;
	}

	/**
	 * This method initializes jPanel3
	 * 
	 * @return javax.swing.JPanel
	 */
	private JPanel getJPanel3() {
		if (jPanel3 == null) {
			jPanel3 = new JPanel();
			jPanel3.setLayout(null);
			jPanel3.setBounds(11, 444, 400, 54);
			jPanel3.setBorder(javax.swing.BorderFactory.createTitledBorder(
					null, "DataSet Content",
					javax.swing.border.TitledBorder.DEFAULT_JUSTIFICATION,
					javax.swing.border.TitledBorder.DEFAULT_POSITION, null,
					null));
			jPanel3.add(getSetRandomVoxelIntensitiesButton(), null);
		}
		return jPanel3;
	}

	/**
	 * This method initializes setRandomVoxelIntensitiesButton
	 * 
	 * @return javax.swing.JButton
	 */
	private JButton getSetRandomVoxelIntensitiesButton() {
		if (setRandomVoxelIntensitiesButton == null) {
			setRandomVoxelIntensitiesButton = new JButton();
			setRandomVoxelIntensitiesButton.setBounds(19, 21, 364, 20);
			setRandomVoxelIntensitiesButton
					.setText("Set Random Voxel Intensities");
			setRandomVoxelIntensitiesButton.setBorder(javax.swing.BorderFactory
					.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
			setRandomVoxelIntensitiesButton
					.addActionListener(new java.awt.event.ActionListener() {
						public void actionPerformed(java.awt.event.ActionEvent e) {
							setRandomIntensities();
						}
					});
		}
		return setRandomVoxelIntensitiesButton;
	}

	public static void main(String[] args) {
		ImarisRemoteControl view = new ImarisRemoteControl();
		view.show();
	}

	private void setRandomIntensities() {
		ActiveXComponent dataSet = imarisApplication
				.getPropertyAsComponent("mDataSet");
		int xSize = dataSet.getPropertyAsInt("mSizeX");
		int ySize = dataSet.getPropertyAsInt("mSizeY");
		int zSize = dataSet.getPropertyAsInt("mSizeZ");
		int chSize = dataSet.getPropertyAsInt("mSizeC");
		int tSize = dataSet.getPropertyAsInt("mSizeT");
		int type = dataSet.getPropertyAsInt("mType");
		for (int z = 0; z < zSize; z++) {
			for (int c = 0; c < chSize; c++) {
				for (int t = 0; t < tSize; t++) {
					SafeArray slice;
					slice = createEmptySliceFor(xSize, ySize, type);
					fillSliceWithRandomValues(xSize, ySize, type, slice);
					Variant vz = new Variant(z);
					Variant vc = new Variant(c);
					Variant vt = new Variant(t);
					Variant vdata = new Variant(slice);
					dataSet.invoke("SetDataSlice", vdata, vz, vc, vt);
				}
			}
		}
	}

	/**
	 * @param xSize
	 * @param ySize
	 * @param type
	 * @param slice
	 */
	private void fillSliceWithRandomValues(int xSize, int ySize, int type, SafeArray slice) {
		Random randomStream = new Random();
		int maxByte = 255;
		int maxShort = 65535;
		for (int x = 0; x < xSize; x++) {
			for (int y = 0; y < ySize; y++) {
				if (type == 2) {
					slice.setChar(x, y, (char) ( randomStream.nextInt(maxShort)));
				}
				if (type == 3) {
					slice
							.setFloat(x, y, (randomStream
									.nextFloat()));
				}
				if (type != 2 && type != 3) {
					slice.setByte(x, y, ((byte) randomStream
							.nextInt(maxByte)));
				}
			}
		}
	}

	/**
	 * @param xSize
	 * @param ySize
	 * @param type
	 * @return
	 */
	private SafeArray createEmptySliceFor(int xSize, int ySize, int type) {
		SafeArray slice;
		slice = new SafeArray(Variant.VariantByte, xSize, ySize);
		if (type == 2) {
			slice = new SafeArray(18, xSize, ySize);
		}
		if (type == 3) {
			slice = new SafeArray(Variant.VariantFloat, xSize, ySize);
		}
		return slice;
	}
} //  @jve:decl-index=0:visual-constraint="10,10"
