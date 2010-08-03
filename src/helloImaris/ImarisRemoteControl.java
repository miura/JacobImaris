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

import ij.IJ;
import ij.ImagePlus;
import ij.ImageStack;
import ij.WindowManager;
import ij.gui.GenericDialog;
import ij.measure.Calibration;
import ij.plugin.CanvasResizer;
import ij.plugin.Duplicator;
import ij.process.ByteProcessor;
import ij.process.FloatProcessor;
import ij.process.ImageProcessor;
import ij.process.ShortProcessor;

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

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Rectangle;
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

	private JButton jButton = null;

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
		this.setSize(440, 577);
		this.setTitle("Imaris interface");
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
			jPanel.add(getJButtonExport3Dstack(), null);
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
					.setText("C:\\Program Files\\Bitplane\\Imaris 7.0.0\\images\\retina.ims");
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
//							xTextField.setText(dataSet
//									.getPropertyAsString("mSizeX"));
//							yTextField.setText(dataSet
//									.getPropertyAsString("mSizeY"));
//							zTextField.setText(dataSet
//									.getPropertyAsString("mSizeZ"));
//							chTextField.setText(dataSet
//									.getPropertyAsString("mSizeC"));
//							tTextField.setText(dataSet
//									.getPropertyAsString("mSizeT"));
							//test 1 
							xTextField.setText(dataSet
									.getProperty("mSizeX").changeType(Variant.VariantString).getString());
							yTextField.setText(dataSet
									.getProperty("mSizeY").changeType(Variant.VariantString).getString());
							zTextField.setText(dataSet
									.getProperty("mSizeZ").changeType(Variant.VariantString).getString());
							chTextField.setText(dataSet
									.getProperty("mSizeC").changeType(Variant.VariantString).getString());
							tTextField.setText(dataSet
									.getProperty("mSizeT").changeType(Variant.VariantString).getString());
							//System.out.println(dataSet.getProperty("mSizeX").changeType(Variant.VariantString));
							//System.out.println("bitdepth " + type);
							
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

	/**
	 * This method initializes jButton	
	 * 	
	 * @return javax.swing.JButton	
	 * @author Kota Miura
	 */
	private JButton getJButtonExport3Dstack() {
		if (jButton == null) {
			jButton = new JButton();
			jButton.setBounds(new Rectangle(33, 510, 363, 26));
			jButton.setText("Export Current Stack to Imaris");
			jButton.addActionListener(new java.awt.event.ActionListener() {
				public void actionPerformed(java.awt.event.ActionEvent e) {
					//System.out.println("actionPerformed()"); 
					ImagePlus imp = WindowManager.getCurrentImage();
					if (imp.getNSlices() == 1) {
						IJ.error("This stack does not have z-depth. Maybe set the image properties");
					} else if (imp.getNFrames()== 1){ 
							if (imp.getNChannels()== 1) 
								ExportDataSetToImaris(1, 1, true);
							else
								IJ.error("Multichannel image not implemented");
					} else {//Nslices > 1 and Nframes > 1		
						if (! imp.isHyperStack())
							IJ.error("Please conver the stack to Hyper Stack");
						else {
							if (imp.getNChannels()== 1) {
								int startSlice = 0;
								int endslice = 1;
								for (int i = 0; i< imp.getNFrames(); i++){
									//imp.setSlice(i * imp.getNSlices() + 1);
									Duplicator dup = new Duplicator();
									ImagePlus imp2 = dup.run(imp, (i * imp.getNSlices() + 1), ((i+1) * imp.getNSlices()));
									//Extractfrom4D ext3D = new Extractfrom4D();
									//ImagePlus imp3D = ext3D.core(imp, 3);
									//imp3D.show();
									if (i==0) 
										ExportDataSetToImaris(1, i+1, true, imp2);
									else
										if (ExportDataSetToImaris(1, i+1, false, imp2)) 
											IJ.log("Time Point " + Integer.toString(i+1)+ " exported.");
									//imp3D.close();
									//imp3D = null;
								}
								
							} else
								IJ.error("Multichannel image not implemented");
						}
					}	
				}
			});
		}
		return jButton;
	}

	public static void main(String[] args) {
		ImarisRemoteControl view = new ImarisRemoteControl();
		//view.show();
		view.setVisible(true);
	}

	private void setRandomIntensities() {
		ActiveXComponent dataSet = imarisApplication
				.getPropertyAsComponent("mDataSet");
		int xSize = dataSet.getProperty("mSizeX").changeType(Variant.VariantInt).getInt();
		int ySize = dataSet.getProperty("mSizeY").changeType(Variant.VariantInt).getInt();
		int zSize = dataSet.getProperty("mSizeZ").changeType(Variant.VariantInt).getInt();
		int chSize = dataSet.getProperty("mSizeC").changeType(Variant.VariantInt).getInt();
		int tSize = dataSet.getProperty("mSizeT").changeType(Variant.VariantInt).getInt();
		int type =  dataSet.getProperty("mType").changeType(Variant.VariantInt).getInt();
		
//		int xSize = (int) dataSet.getProperty("mSizeX").changeType(Variant.VariantLongInt).getLong();
//		int ySize = (int) dataSet.getProperty("mSizeY").changeType(Variant.VariantLongInt).getLong();
//		int zSize = (int) dataSet.getProperty("mSizeZ").changeType(Variant.VariantLongInt).getLong();
//		int chSize = (int) dataSet.getProperty("mSizeC").changeType(Variant.VariantLongInt).getLong();
//		int tSize = (int) dataSet.getProperty("mSizeT").changeType(Variant.VariantLongInt).getLong();
//		int type = (int) dataSet.getProperty("mType").changeType(Variant.VariantLongInt).getLong();
		
		//int type = dataSet.getPropertyAsInt("mType");

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
					slice.setFloat(x, y, (randomStream.nextFloat()));
				}
				if (type != 2 && type != 3) {
					slice.setByte(x, y, ((byte) randomStream.nextInt(maxByte)));
				}
			}
		}
	}
	
	/**
	 * @author Kota Miura
	 */
	private void fillSliceWithStackSlice(int xSize, int ySize, int type, SafeArray slice, ImageProcessor ip) {

		for (int x = 0; x < xSize; x++) {
			for (int y = 0; y < ySize; y++) {
				if (type == 2) {
					//slice.setChar(x, y, (char) ( randomStream.nextInt(maxShort)));
					slice.setChar(x, y, (char) ip.get(x,y));
				}
				if (type == 3) {
					slice.setFloat(x, y, ip.getPixelValue(x,y));
				}
				if (type != 2 && type != 3) {
					slice.setByte(x, y, ((byte) ip.get(x,y)));
				}
			}
		}
	}

	private void fillVolumeWithStack(int xSize, int ySize, int zSize, int type, 
			SafeArray stack, ImageStack ims) {

		int[] xyzCoord = new int[3];
		int pixvalue = 0;
		int currentpos = 0;
		byte[] pixelsByte = new byte[xSize * ySize];
		char[] pixelsChar = new char[xSize * ySize];
		float[] pixels  = new float[xSize * ySize];		
		for (int z = 0; z < zSize; z++) {
			if (type == 1) pixelsByte = (byte[]) ims.getProcessor(z+1).getPixels();	
			if (type == 2) pixelsChar = (char[]) ims.getProcessor(z+1).getPixels();	
			if (type == 3) pixels = (float[]) ims.getProcessor(z+1).getPixels();				
			for (int y = 0; y < ySize; y++) {
				for (int x = 0; x < xSize; x++) {
					currentpos = xSize * y + x; 
					xyzCoord[0] = x;
					xyzCoord[1] = y;
					xyzCoord[2] = z;
					if (type != 2 && type != 3)
						stack.setByte(xyzCoord, (byte) pixelsByte[currentpos]);
					if (type == 2) 
						stack.setChar(xyzCoord, (char) pixelsChar[currentpos]);
					if (type == 3) 
						stack.setFloat(xyzCoord, (float) pixelsChar[currentpos]);

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
	/**
	 * @author Kota Miura
	 * @param xSize
	 * @param ySize
	 * @param zSize
	 * @param type
	 * @return
	 */
	private SafeArray createEmptyVolumeFor(int xSize, int ySize, int zSize, int type) {
		SafeArray volume;
		int[] lowbounds = {0,0,0};
		int[] upbounds = {xSize, ySize, zSize};
		
		volume = new SafeArray(Variant.VariantByte, lowbounds, upbounds);
		if (type == 2) {
			volume = new SafeArray(18, lowbounds, upbounds);
		}
		if (type == 3) {
			volume = new SafeArray(Variant.VariantFloat, lowbounds, upbounds);
		}
		return volume;
	}
	
	public boolean ExportDataSetToImaris(int paramInt1, int paramInt2, boolean paramBoolean){
		boolean done = false;
		ImagePlus localImagePlus = IJ.getImage();
		done = ExportDataSetToImaris(paramInt1, paramInt2, paramBoolean, 
				localImagePlus);
		return done;
	}
			
	//modified bpImaris_Adaptor2 for use with JACOB-imaris by Volker. 
	//implementation of export button action
	//TODO Geometry (voxel size) should be controllable from ImageJ
	//	--> for this, use set mExtendMinX, mExtendMaxX pair. 
	public boolean ExportDataSetToImaris(int paramInt1, int paramInt2, boolean paramBoolean, 
			ImagePlus localImagePlus) {
		if (imarisApplication != null) {
			imarisApplication.setProperty("mUserControl", true); //this keeps Imaris Opened
			int ijCh = paramInt1;	//channel
			int ijT = paramInt2;	//time point
 
			boolean bool = paramBoolean;	//true if newly create a stack in Imaris. 
			//ImagePlus localImagePlus = IJ.getImage();
			localImagePlus.lock();
			if (localImagePlus == null) {
				IJ.showMessage("Please select an ImageJ image before exporting.");
				return false; 
			} 
			int ijbitdep = localImagePlus.getBitDepth();
			ImageStack localImageStack = localImagePlus.getStack();
			int ijXsize = localImageStack.getWidth();
			int ijYsize = localImageStack.getHeight();
			int ijnSlices = localImageStack.getSize();	//total number of slices, regardless of t or c
			int i6 = 1;
			int i7 = 1;
			int ijZsize = localImagePlus.getNSlices();
			int ijTsize = localImagePlus.getNFrames();
			int ijCsize = localImagePlus.getNChannels();
			
			// IDataSet localIDataSet = null;
			ActiveXComponent localIDataSet = null;
//			Object localObject1;
//			Object localObject2;
			Object localObject3;
			Object localObject4;
//			Object localObject5;
			Object localObject6;
			int i1;
			int i3;
			int i5;
			//try { localIDataSet = this.mImaris.getMDataSet();
			try { localIDataSet = imarisApplication.getPropertyAsComponent("mDataSet");
				//((Object) localIDataSet).setAutoDelete(false);
			
				//create new object in Imars
				if ((localIDataSet == null) || (bool == true)) {
					Variant[] parameter = new Variant[6];	//In JACOB, this contains dimentional info and will be passed with 'invoke'
					parameter[1] = new Variant(ijXsize);	//x
					parameter[2] = new Variant(ijYsize); //y
					parameter[3] = new Variant(ijnSlices); //z
					parameter[4] = new Variant(i6); //ch
					parameter[5] = new Variant(i7); //t
					if (ijbitdep == 8)	{	//8-bit image
						parameter[0] = new Variant(1);	//type 8bit
						localIDataSet.invoke("Create", parameter);		
						//localIDataSet.create(new tType_(1L), new UInt32(l), new UInt32(i2), new UInt32(i4), new UInt32(i6), new UInt32(i7));
					} else if (ijbitdep == 16) {
							parameter[0] = new Variant(2);	//type 16 bit
							localIDataSet.invoke("Create", parameter);				
							//localIDataSet.create(new tType_(2L), new UInt32(l), new UInt32(i2), new UInt32(i4), new UInt32(i6), new UInt32(i7));
					} else if (ijbitdep == 32) {
								parameter[0] = new Variant(3);	//type 32 bit
								localIDataSet.invoke("Create", parameter);						
								//localIDataSet.create(new tType_(3L), new UInt32(l), new UInt32(i2), new UInt32(i4), new UInt32(i6), new UInt32(i7)); 
					} else {
						if (ijbitdep == 24) {
							IJ.showMessage("Error", "Please use an image of type 8, 16 or float.");
							localImagePlus.unlock();
							//localIDataSet.release();
							localIDataSet.safeRelease();
							localIDataSet = null;
							return false;
						}
						IJ.showMessage("Error", "Unknown image type");
						localImagePlus.unlock();
						//localIDataSet.release();
						localIDataSet.safeRelease();
						localIDataSet = null;
						return false;
					}
					ijCh = 1;
					ijT = 1;
				//append objects to the current object already in Imaris	
				//TODO: implement resize protocol. 
				} else { 
					//UInt32 localUInt32 = localIDataSet.getMSizeX();	//TODO I don't understand why specificallt this xsize is UInt32.
					//localObject1 = localIDataSet.getMSizeY();
					//localObject2 = localIDataSet.getMSizeZ();
					//localObject3 = localIDataSet.getMSizeC();
					//localObject4 = localIDataSet.getMSizeT();
					//localObject5 = localIDataSet.getMType();
					Variant[] resizepara = new Variant[10];
					resizepara[0] = new Variant(0);
					//resizepara[1] = new Variant(localIDataSet.getProperty("mSizeX"));
					resizepara[1] = localIDataSet.getProperty("mSizeX");					
					resizepara[2] = new Variant(0);
					resizepara[3] = localIDataSet.getProperty("mSizeY");
					resizepara[4] = new Variant(0);
					resizepara[5] = localIDataSet.getProperty("mSizeZ");
					resizepara[6] = new Variant(0);
					resizepara[7] = localIDataSet.getProperty("mSizeC");
					resizepara[8] = new Variant(0);
					resizepara[9] = localIDataSet.getProperty("mSizeT");
					int imXsize = resizepara[1].changeType(Variant.VariantInt).getInt();
					int imYsize = resizepara[3].changeType(Variant.VariantInt).getInt();
					int imZsize = resizepara[5].changeType(Variant.VariantInt).getInt();
					int imCsize = resizepara[7].changeType(Variant.VariantInt).getInt();					
					int imTsize = resizepara[9].changeType(Variant.VariantInt).getInt();
					
					int imtype = localIDataSet.getPropertyAsInt("mType");	//bit depth: 0 = unknown, 1 = 8bit. 2 = 16bit, 3 = 32bit float
					
					// check image size and reject 
					//if (((k == 8) && (((tType_)localObject5).getValue() != 1L)) || ((k == 16) && (((tType_)localObject5).getValue() != 2L)) || ((k == 32) && (((tType_)localObject5).getValue() != 3L)))
					if (((ijbitdep == 8) && (imtype != 1)) || ((ijbitdep == 16) && (imtype != 2)) || ((ijbitdep == 32) && (imtype != 3)))
					{
						IJ.showMessage("Error", "The type of the ImageJ DataSet is not the same as Imaris DataSet!");
						localImagePlus.unlock();
						//localIDataSet.release();
						localIDataSet = null;
						return false; 
					}
					if (ijbitdep == 24) {
						IJ.showMessage("Error", "ImageJ RGB type is not yet usable. Please convert to 8,16 or 32 bit image.");
						localImagePlus.unlock();
						//localIDataSet.release();
						localIDataSet = null;
						return false;
					}
					// check image size and resize. Expand to fit larger of either IJimage or imImage. 
					//TODO these resizing are too crude. should eliminate or fix. 
					// e.g. resizing origin point is always (0, 0)
					//if (((int)localUInt32.getValue() != xsize) || ((int)((UInt32)localObject1).getValue() != ysize) || ((int)((UInt32)localObject2).getValue() != nSlices))
					if ((imXsize != ijXsize) 
							|| (imYsize != ijYsize) 
								|| (imZsize != ijnSlices))
					{
						if (imXsize < ijXsize) {
							//localIDataSet.resize(new Int32(0), new UInt32(xsize), new Int32(0), (UInt32)localObject1, new Int32(0), (UInt32)localObject2, new Int32(0), (UInt32)localObject3, new Int32(0), (UInt32)localObject4);
							resizepara[1] = new Variant(ijXsize);
							localIDataSet.invoke("Resize", resizepara);
							//localUInt32.setValue(xsize);
							imXsize = ijXsize;
						}
						else if (imXsize > ijXsize) {
							localObject6 = new CanvasResizer();
							//localImageStack = ((CanvasResizer)localObject6).expandStack(localImageStack, (int)localUInt32.getValue(), ysize, 0, 0);
							localImageStack = ((CanvasResizer)localObject6).expandStack(localImageStack, imXsize, ijYsize, 0, 0);
							//i1 = (int)localUInt32.getValue();
							//i1 = imXsize;
							ijXsize = imXsize;
						}
 
						if (imYsize < ijYsize) {
							//localIDataSet.resize(new Int32(0), localUInt32, new Int32(0), new UInt32(ysize), new Int32(0), (UInt32)localObject2, new Int32(0), (UInt32)localObject3, new Int32(0), (UInt32)localObject4);
							resizepara[3] = new Variant(ijYsize);
							localIDataSet.invoke("Resize", resizepara);
							//((UInt32)localObject1).setValue(ysize);
							imYsize = ijYsize;
						}
						else if (imYsize > ijYsize) {
							localObject6 = new CanvasResizer();
							//localImageStack = ((CanvasResizer)localObject6).expandStack(localImageStack, i1, (int)((UInt32)localObject1).getValue(), 0, 0);
							localImageStack = ((CanvasResizer)localObject6).expandStack(localImageStack, ijXsize, imYsize, 0, 0);
							//i3 = imYsize;//(int)((UInt32)localObject1).getValue();
							ijYsize = imXsize;
						}

						if (imZsize < ijnSlices) {
							//localIDataSet.resize(new Int32(0), localUInt32, new Int32(0), 
								//(UInt32)localObject1, new Int32(0), new UInt32(nSlices), new Int32(0), 
									//(UInt32)localObject3, new Int32(0), (UInt32)localObject4);
							resizepara[5] = new Variant(ijnSlices);
							localIDataSet.invoke("Resize", resizepara);							
							imZsize = ijnSlices;//((UInt32)localObject2).setValue(nSlices);
						}
	
						else if (imZsize > ijnSlices) {
							if (ijbitdep == 8) {
								localObject6 = new ByteProcessor(ijXsize, ijYsize);
							} else if (ijbitdep == 16) {
								localObject6 = new ShortProcessor(ijXsize, ijYsize);
							} else if (ijbitdep == 32) {
								localObject6 = new FloatProcessor(ijXsize, ijYsize);
							} else {
								localImagePlus.unlock();
								//localIDataSet.release();
								localIDataSet = null;
								return false;
							}
							((ImageProcessor)localObject6).setColor(new Color(0.0F, 0.0F, 0.0F));
							//for (int i11 = 0; i11 < (int)((UInt32)localObject2).getValue() - nSlices; ++i11) {
							for (int i11 = 0; i11 < imZsize - ijnSlices; ++i11) {
								localImageStack.addSlice("", (ImageProcessor)localObject6);
							}
							//i5 = imZsize;//(int)((UInt32)localObject2).getValue();
							ijnSlices = imZsize;
						}
					}
					//channel exceeds
					//if (i > (int)((UInt32)localObject3).getValue()) {
					if (ijCh > imCsize) {
						//localObject3 = new UInt32(((UInt32)localObject3).getValue() + 1L);
						localObject3 = new Variant(imCsize + 1);
						//i = (int)((UInt32)localObject3).getValue();
						//ijCh = imCsize;
						//localIDataSet.resize(new Int32(0), localUInt32, new Int32(0), (UInt32)localObject1, new Int32(0), (UInt32)localObject2, new Int32(0), new UInt32(i), new Int32(0), (UInt32)localObject4);
						resizepara[7] = new Variant(ijCh);
						localIDataSet.invoke("Resize", resizepara);
						imCsize = ijCh;
					}

					//if (j > (int)((UInt32)localObject4).getValue()) {
					if (ijT > imTsize) {
						//localObject4 = new UInt32(((UInt32)localObject4).getValue() + 1L);
						localObject4 = new Variant(imTsize + 1);
						//j = (int)((UInt32)localObject4).getValue();
						//ijT = imTsize;
						//localIDataSet.resize(new Int32(0), localUInt32, new Int32(0), (UInt32)localObject1, new Int32(0), (UInt32)localObject2, new Int32(0), (UInt32)localObject3, new Int32(0), new UInt32(j));
						resizepara[9] = new Variant(ijT);
						localIDataSet.invoke("Resize", resizepara);
						imTsize = ijT;
					}
				}
				//localIDataSet.release();
				localIDataSet = null;
	
			} catch (Exception localException1) {

				localException1.printStackTrace();
				IJ.showMessage("Error", "Unknown Exception occured");
				localImagePlus.unlock();
				//localIDataSet.release();
				localIDataSet = null;
				return false;
			}
					// Send the data to Imaris...
			try {
				//localIDataSet = this.mImaris.getMDataSet();
				localIDataSet = imarisApplication.getPropertyAsComponent("mDataSet");
				//localIDataSet.setAutoDelete(false);
				if ((ijbitdep == 8) || (ijbitdep == 16) || (ijbitdep == 32)){				
					int imtype = 0;
					if (ijbitdep == 8) imtype = 1;
					if (ijbitdep == 16) imtype = 2;
					if (ijbitdep == 32) imtype = 3;
					TranferScaleIJ2Imaris(localIDataSet, localImagePlus);	//sets scale in Imaris ([Edit-> ImageProperties -> geometry])
/* old version, slice by slice
					for (int i8 = 0; i8 < ijnSlices; ++i8) {
						localObject1 = localImageStack.getProcessor(i8 + 1);
						((ImageProcessor)localObject1).flipVertical();
	
						
						SafeArray slice;
						slice = createEmptySliceFor(ijXsize, ijYsize, imtype);
						fillSliceWithStackSlice(ijXsize, ijYsize, imtype, slice, (ImageProcessor) localObject1);
						Variant vz = new Variant(i8);
						Variant vc = new Variant(ijCh-1);
						Variant vt = new Variant(ijT-1);
						Variant vdata = new Variant(slice);
						localIDataSet.invoke("SetDataSlice", vdata, vz, vc, vt);						
						((ImageProcessor)localObject1).flipVertical();
						IJ.showProgress(i8, localImageStack.getSize() - 1);							
					}
*/					
					SafeArray volume;
					volume = createEmptyVolumeFor(ijXsize, ijYsize, ijnSlices, imtype);
					fillVolumeWithStack(ijXsize, ijYsize, ijnSlices, imtype, volume, localImageStack);
					Variant vc = new Variant(ijCh-1);
					Variant vt = new Variant(ijT-1);
					Variant vdata = new Variant(volume);
					localIDataSet.invoke("SetDataVolume", vdata, vc, vt);					
					
				} else {
					if (ijbitdep == 24) {
						IJ.showMessage("Error", "RGB images cannot be exported now, please use 8,16 or 32-bit images");
						localImagePlus.unlock();
						//localIDataSet.release();
						localIDataSet.safeRelease();
						localIDataSet = null;
						return false;
					}
					IJ.showMessage("Error", "Unknown image type");
					localImagePlus.unlock();
					//localIDataSet.release();
					localIDataSet.safeRelease();
					localIDataSet = null;
					return false;
				}		
			}
			catch (Exception localException2) {
				localException2.printStackTrace();
				IJ.showMessage("Error", "Unknown Exception occured while stack transfer");
				localImagePlus.unlock();
				//localIDataSet.release();
				localIDataSet.safeRelease();
				localIDataSet = null;
				return false;
			}
			localImagePlus.unlock();
			//localIDataSet.release();
			localIDataSet.safeRelease();
			localIDataSet = null;
			return true;
		}
		return false;
	}

	public boolean TranferScaleIJ2Imaris(ActiveXComponent localIDataSet, ImagePlus imp){
		if (imp == null) return false;
		Calibration cal = new Calibration();
		cal = imp.getCalibration();
		float xscale = (float) cal.pixelWidth;
		float yscale = (float) cal.pixelHeight;
		float zscale = (float) cal.pixelDepth;		
		SetVoxelScale(localIDataSet, xscale, yscale, zscale);
		
		return true;
	}
	
	public boolean SetVoxelScale(ActiveXComponent localIDataSet, 
			float xscale, float yscale, float zscale){
		try {
			int imXsize = localIDataSet.getProperty("mSizeX").changeType(Variant.VariantInt).getInt();
			int imYsize = localIDataSet.getProperty("mSizeY").changeType(Variant.VariantInt).getInt();
			int imZsize = localIDataSet.getProperty("mSizeZ").changeType(Variant.VariantInt).getInt();
			float Scaled_imXsize = imXsize * xscale;
			float Scaled_imYsize = imYsize * yscale;
			float Scaled_imZsize = imZsize * zscale;
			localIDataSet.setProperty("mExtendMinX", new Variant(0));
			localIDataSet.setProperty("mExtendMaxX", new Variant(Scaled_imXsize));
			localIDataSet.setProperty("mExtendMinY", new Variant(0));
			localIDataSet.setProperty("mExtendMaxY", new Variant(Scaled_imYsize));
			localIDataSet.setProperty("mExtendMinZ", new Variant(0));
			localIDataSet.setProperty("mExtendMaxZ", new Variant(Scaled_imZsize));

		} catch (Exception localException3) {
			localException3.printStackTrace();
			IJ.showMessage("Error", "Unknown Exception occured while setting scale");
			return false;
		}
		return true;
	}

	public class Extractfrom4D {

		private int current3DstackStart = 0; 
		private int current3DstackEnd = 1; 
		
		public void run(int extractdimension){
			ImagePlus imp = WindowManager.getCurrentImage();
			ImagePlus imp2 = core(imp, extractdimension);
			imp2.show();
		}
			
		public ImagePlus core(ImagePlus imp, int extractdimension){
			if (imp.getNDimensions() < 4) {
				IJ.error("Check Image Properties: this stack has no t-dimension");
				return null;
			}
			Calibration cal = new Calibration(imp);
			int nSlices = imp.getNSlices();
			int nFrames = imp.getNFrames();
			int nChannels = imp.getNChannels();
			
			int newframes = 1; 
			if (extractdimension == 3) {
				getStartAndEnd3D(imp.getFrame(), nSlices, nChannels);
			} else {	//extract a range of 4D stack
				newframes = getStartAndEnd4D(nFrames, nSlices, nChannels);
			}
			
			Duplicator dup = new Duplicator();
			ImagePlus imp2 = dup.run(imp, current3DstackStart, current3DstackEnd);
			if (extractdimension == 4){
				//HyperStackConverter hyp = new HyperStackConverter();
				//hyp.shuffle();
				imp2.setDimensions(nChannels, nSlices, newframes);
				imp2.setOpenAsHyperStack(true);
			}
			return imp2;
		}
		
		void getStartAndEnd3D(int currentFrame, int nSlices, int nChannels){
			current3DstackStart = (currentFrame-1)*nSlices*nChannels +1; 
			current3DstackEnd = currentFrame*nSlices*nChannels; 		
		}
		int getStartAndEnd4D(int nFrames, int nSlices, int nChannels){
			GenericDialog gd = new GenericDialog("Time Frame Range");
			gd.addNumericField("Start t-Frame (>=1) :", 1, 0);
			gd.addNumericField("End t-Frame (<"+ nFrames+") :", nFrames, 0);
			gd.showDialog();
			if (gd.wasCanceled()) 
				return -1;
			int startframe = (int) gd.getNextNumber();
			int endframe = (int) gd.getNextNumber();
			if ((startframe<1) || (endframe>nFrames))
					return -1;
			if (startframe > endframe) return -1;
			current3DstackStart = (startframe-1)*nSlices*nChannels +1;
			current3DstackEnd = endframe*nSlices*nChannels; 
			return (endframe - startframe + 1);
		}
	}	
	
} //  @jve:decl-index=0:visual-constraint="10,10"
