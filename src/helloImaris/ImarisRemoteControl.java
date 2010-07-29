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
import ij.plugin.CanvasResizer;
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
					System.out.println("actionPerformed()"); // TODO Auto-generated Event stub actionPerformed()
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
	
	//modified bpImaris_Adaptor2 for use with jacob-imaris by Volker. 
	//implementation of export button acition
	public boolean ExportDataSetToImaris(int paramInt1, int paramInt2, boolean paramBoolean) {
		if (imarisApplication != null) {
			int i = paramInt1;
	/* 1396 */         int j = paramInt2;
	/*      */ 
	/* 1398 */         boolean bool = paramBoolean;
	/*      */ 
	/* 1402 */         ImagePlus localImagePlus = IJ.getImage();
	/* 1403 */         localImagePlus.lock();
	/* 1404 */         if (localImagePlus == null) {
	/* 1405 */           IJ.showMessage("Please select an ImageJ image before exporting.");
	/* 1406 */           return false; 
						} 
						int k = localImagePlus.getBitDepth();
	/* 1409 */         ImageStack localImageStack = localImagePlus.getStack();
	/* 1410 */         int l = localImageStack.getWidth();
	/* 1411 */         int i2 = localImageStack.getHeight();
	/* 1412 */         int i4 = localImageStack.getSize();
	/* 1413 */         int i6 = 1;
	/* 1414 */         int i7 = 1;
	/*      */ 
	/* 1418 */         // IDataSet localIDataSet = null;
						ActiveXComponent localIDataSet = null;
	/*      */         Object localObject1;
	/*      */         Object localObject2;
	/*      */         Object localObject3;
	/*      */         Object localObject4;
	/*      */         Object localObject5;
	/*      */         Object localObject6;
	/*      */         int i1;
	/*      */         int i3;
	/*      */         int i5;
	/*      */         //try { localIDataSet = this.mImaris.getMDataSet();
						try { localIDataSet = imarisApplication.getPropertyAsComponent("mDataSet");
	/* 1421 */           ((Object) localIDataSet).setAutoDelete(false);

	
						//create new object in Imars
	/* 1423 */           if ((localIDataSet == null) || (bool == true))
	/*      */           {
							Variant[] parameter = new Variant[6];
							parameter[1] = new Variant(l);	//x
							parameter[2] = new Variant(i2); //y
							parameter[3] = new Variant(i4); //z
							parameter[4] = new Variant(i6); //ch
							parameter[5] = new Variant(i7); //t
	/* 1425 */             if (k == 8)
	/*      */             {
							parameter[0] = new Variant(1);	//type 8bit
							localIDataSet.invoke("Create", parameter);		
	/* 1427 */              //localIDataSet.create(new tType_(1L), new UInt32(l), new UInt32(i2), new UInt32(i4), new UInt32(i6), new UInt32(i7));
	/* 1428 */             } else if (k == 16)
	/*      */             {
							parameter[0] = new Variant(2);	//type 16 bit
							localIDataSet.invoke("Create", parameter);				
	/* 1430 */              //localIDataSet.create(new tType_(2L), new UInt32(l), new UInt32(i2), new UInt32(i4), new UInt32(i6), new UInt32(i7));
	/* 1431 */             } else if (k == 32)
	/*      */             {
							parameter[0] = new Variant(3);	//type 32 bit
							localIDataSet.invoke("Create", parameter);						
	/* 1433 */              //localIDataSet.create(new tType_(3L), new UInt32(l), new UInt32(i2), new UInt32(i4), new UInt32(i6), new UInt32(i7)); } else {
	/* 1434 */               if (k == 24)
	/*      */               {
	/* 1437 */                 IJ.showMessage("Error", "Please use an image of type 8, 16 or float.");
	/* 1438 */                 localImagePlus.unlock();
	/* 1439 */                 localIDataSet.release();
	/* 1440 */                 localIDataSet = null;
	/* 1441 */                 return false;
	/*      */               }
	/* 1443 */               IJ.showMessage("Error", "Unknown image type");
	/* 1444 */               localImagePlus.unlock();
	/* 1445 */               localIDataSet.release();
	/* 1446 */               localIDataSet = null;
	/* 1447 */               return false;
	/*      */             }
	/*      */ 
	/* 1451 */             i = 1;
	/* 1452 */             j = 1;
	/*      */           }
	/*      */           else
							//append objects to the current already in Imaris
	/*      */           {
	/* 1457 */             UInt32 localUInt32 = localIDataSet.getMSizeX();
	/* 1458 */             localObject1 = localIDataSet.getMSizeY();
	/* 1459 */             localObject2 = localIDataSet.getMSizeZ();
	/* 1460 */             localObject3 = localIDataSet.getMSizeC();
	/* 1461 */             localObject4 = localIDataSet.getMSizeT();
	/* 1462 */             localObject5 = localIDataSet.getMType();
	/*      */ 
	/* 1465 */             if (((k == 8) && (((tType_)localObject5).getValue() != 1L)) || ((k == 16) && (((tType_)localObject5).getValue() != 2L)) || ((k == 32) && (((tType_)localObject5).getValue() != 3L)))
	/*      */             {
	/* 1468 */               IJ.showMessage("Error", "The type of the ImageJ DataSet is not the same as Imaris DataSet!");
	/* 1469 */               localImagePlus.unlock();
	/* 1470 */               localIDataSet.release();
	/* 1471 */               localIDataSet = null;
	/* 1472 */               return false; }
	/* 1473 */             if (k == 24) {
	/* 1474 */               IJ.showMessage("Error", "ImageJ RGB type is not yet usable. Please convert to 8,16 or 32 bit image.");
	/* 1475 */               localImagePlus.unlock();
	/* 1476 */               localIDataSet.release();
	/* 1477 */               localIDataSet = null;
	/* 1478 */               return false;
	/*      */             }
	/*      */ 
	/* 1483 */             if (((int)localUInt32.getValue() != l) || ((int)((UInt32)localObject1).getValue() != i2) || ((int)((UInt32)localObject2).getValue() != i4))
	/*      */             {
	/* 1485 */               if ((int)localUInt32.getValue() < l) {
	/* 1486 */                 localIDataSet.resize(new Int32(0), new UInt32(l), new Int32(0), (UInt32)localObject1, new Int32(0), (UInt32)localObject2, new Int32(0), (UInt32)localObject3, new Int32(0), (UInt32)localObject4);
	/* 1487 */                 localUInt32.setValue(l);
	/*      */               }
	/* 1490 */               else if ((int)localUInt32.getValue() > l) {
	/* 1491 */                 localObject6 = new CanvasResizer();
	/* 1492 */                 localImageStack = ((CanvasResizer)localObject6).expandStack(localImageStack, (int)localUInt32.getValue(), i2, 0, 0);
	/* 1493 */                 i1 = (int)localUInt32.getValue();
	/*      */               }
	/*      */ 
	/* 1497 */               if ((int)((UInt32)localObject1).getValue() < i2) {
	/* 1498 */                 localIDataSet.resize(new Int32(0), localUInt32, new Int32(0), new UInt32(i2), new Int32(0), (UInt32)localObject2, new Int32(0), (UInt32)localObject3, new Int32(0), (UInt32)localObject4);
	/* 1499 */                 ((UInt32)localObject1).setValue(i2);
	/*      */               }
	/* 1502 */               else if ((int)((UInt32)localObject1).getValue() > i2) {
	/* 1503 */                 localObject6 = new CanvasResizer();
	/* 1504 */                 localImageStack = ((CanvasResizer)localObject6).expandStack(localImageStack, i1, (int)((UInt32)localObject1).getValue(), 0, 0);
	/* 1505 */                 i3 = (int)((UInt32)localObject1).getValue();
	/*      */               }
	/*      */ 
	/* 1509 */               if ((int)((UInt32)localObject2).getValue() < i4) {
	/* 1510 */                 localIDataSet.resize(new Int32(0), localUInt32, new Int32(0), (UInt32)localObject1, new Int32(0), new UInt32(i4), new Int32(0), (UInt32)localObject3, new Int32(0), (UInt32)localObject4);
	/* 1511 */                 ((UInt32)localObject2).setValue(i4);
	/*      */               }
	/* 1514 */               else if ((int)((UInt32)localObject2).getValue() > i4)
	/*      */               {
	/* 1516 */                 if (k == 8) {
	/* 1517 */                   localObject6 = new ByteProcessor(i1, i3);
	/* 1518 */                 } else if (k == 16) {
	/* 1519 */                   localObject6 = new ShortProcessor(i1, i3);
	/* 1520 */                 } else if (k == 32) {
	/* 1521 */                   localObject6 = new FloatProcessor(i1, i3);
	/*      */                 } else {
	/* 1523 */                   localImagePlus.unlock();
	/* 1524 */                   localIDataSet.release();
	/* 1525 */                   localIDataSet = null;
	/* 1526 */                   return false;
	/*      */                 }
	/* 1528 */                 ((ImageProcessor)localObject6).setColor(new Color(0.0F, 0.0F, 0.0F));
	/* 1529 */                 for (int i11 = 0; i11 < (int)((UInt32)localObject2).getValue() - i4; ++i11) {
	/* 1530 */                   localImageStack.addSlice("", (ImageProcessor)localObject6);
	/*      */                 }
	/* 1532 */                 i5 = (int)((UInt32)localObject2).getValue();
	/*      */               }
	/*      */ 
	/*      */             }
	/*      */ 
	/* 1538 */             if (i > (int)((UInt32)localObject3).getValue()) {
	/* 1539 */               localObject3 = new UInt32(((UInt32)localObject3).getValue() + 1L);
	/* 1540 */               i = (int)((UInt32)localObject3).getValue();
	/* 1541 */               localIDataSet.resize(new Int32(0), localUInt32, new Int32(0), (UInt32)localObject1, new Int32(0), (UInt32)localObject2, new Int32(0), new UInt32(i), new Int32(0), (UInt32)localObject4);
	/*      */             }
	/*      */ 
	/* 1544 */             if (j > (int)((UInt32)localObject4).getValue()) {
	/* 1545 */               localObject4 = new UInt32(((UInt32)localObject4).getValue() + 1L);
	/* 1546 */               j = (int)((UInt32)localObject4).getValue();
	/* 1547 */               localIDataSet.resize(new Int32(0), localUInt32, new Int32(0), (UInt32)localObject1, new Int32(0), (UInt32)localObject2, new Int32(0), (UInt32)localObject3, new Int32(0), new UInt32(j));
	/*      */             }
	/*      */           }
	/* 1550 */           localIDataSet.release();
	/* 1551 */           localIDataSet = null;
	/*      */         } catch (Exception localException1) {
	/* 1553 */           localException1.printStackTrace();
	/* 1554 */           IJ.showMessage("Error", "Unknown Exception occured");
	/* 1555 */           localImagePlus.unlock();
	/* 1556 */           localIDataSet.release();
	/* 1557 */           localIDataSet = null;
	/* 1558 */           return false;
	/*      */         }
	/*      */ 
	/*      */         try
	/*      */         {
	/* 1564 */           localIDataSet = this.mImaris.getMDataSet();
	/* 1565 */           localIDataSet.setAutoDelete(false);
	/*      */ 
	/* 1567 */           for (int i8 = 0; i8 < i5; ++i8)
	/*      */           {
	/* 1569 */             localObject1 = localImageStack.getProcessor(i8 + 1);
	/* 1570 */             ((ImageProcessor)localObject1).flipVertical();
	/*      */ 
	/* 1572 */             if (k == 8) {
	/* 1573 */               localObject2 = (byte[])localImageStack.getPixels(i8 + 1);
	/*      */ 
	/* 1575 */               localObject3 = new SafeArrayBound[] { new SafeArrayBound(i1, 0), new SafeArrayBound(i3, 0) };
	/*      */ 
	/* 1577 */               localObject4 = new SafeArray(localObject3, localObject2);
	/*      */ 
	/* 1579 */               localObject5 = new Variant((SafeArray)localObject4);
	/* 1580 */               localIDataSet.setDataSlice((Variant)localObject5, new UInt32(i8), new UInt32(i - 1), new UInt32(j - 1));
	/*      */             }
	/* 1582 */             else if (k == 16) {
	/* 1583 */               localObject2 = (short[])localImageStack.getPixels(i8 + 1);
	/*      */ 
	/* 1585 */               localObject3 = new SafeArrayBound[] { new SafeArrayBound(i1, 0), new SafeArrayBound(i3, 0) };
	/*      */ 
	/* 1587 */               localObject4 = new byte[2 * localObject2.length];
	/* 1588 */               for (int i9 = 0; i9 < localObject2.length; ++i9) {
	/* 1589 */                 localObject4[(2 * i9)] = (byte)(0xFF & localObject2[i9]);
	/* 1590 */                 localObject4[(2 * i9 + 1)] = (byte)(0xFF & localObject2[i9] >> 8);
	/*      */               }
	/* 1592 */               SafeArray localSafeArray1 = new SafeArray(localObject3, localObject4, (bpImaris_Adapter2.class$com$jniwrapper$UInt16 == null) ? (bpImaris_Adapter2.class$com$jniwrapper$UInt16 = bpImaris_Adapter2.class$("com.jniwrapper.UInt16")) : bpImaris_Adapter2.class$com$jniwrapper$UInt16);
	/* 1593 */               localObject6 = new Variant(localSafeArray1);
	/* 1594 */               localIDataSet.setDataSlice((Variant)localObject6, new UInt32(i8), new UInt32(i - 1), new UInt32(j - 1));
	/*      */             }
	/* 1596 */             else if (k == 32) {
	/* 1597 */               localObject2 = (float[])localImageStack.getPixels(i8 + 1);
	/*      */ 
	/* 1599 */               localObject3 = new SafeArrayBound[] { new SafeArrayBound(i1, 0), new SafeArrayBound(i3, 0) };
	/*      */ 
	/* 1602 */               localObject4 = new byte[4 * localObject2.length];
	/* 1603 */               for (int i10 = 0; i10 < localObject2.length; ++i10) {
	/* 1604 */                 localObject4[(4 * i10)] = (byte)(0xFF & Float.floatToIntBits(localObject2[i10]));
	/* 1605 */                 localObject4[(4 * i10 + 1)] = (byte)(0xFF & Float.floatToIntBits(localObject2[i10]) >> 8);
	/* 1606 */                 localObject4[(4 * i10 + 2)] = (byte)(0xFF & Float.floatToIntBits(localObject2[i10]) >> 16);
	/* 1607 */                 localObject4[(4 * i10 + 3)] = (byte)(0xFF & Float.floatToIntBits(localObject2[i10]) >> 24);
	/*      */               }
	/* 1609 */               SafeArray localSafeArray2 = new SafeArray(localObject3, localObject4, (bpImaris_Adapter2.class$com$jniwrapper$SingleFloat == null) ? (bpImaris_Adapter2.class$com$jniwrapper$SingleFloat = bpImaris_Adapter2.class$("com.jniwrapper.SingleFloat")) : bpImaris_Adapter2.class$com$jniwrapper$SingleFloat);
	/* 1610 */               localObject6 = new Variant(localSafeArray2);
	/* 1611 */               localIDataSet.setDataSlice((Variant)localObject6, new UInt32(i8), new UInt32(i - 1), new UInt32(j - 1));
	/*      */             } else {
	/* 1613 */               if (k == 24) {
	/* 1614 */                 IJ.showMessage("Error", "RGB images cannot be exported now, please use 8,16 or 32-bit images");
	/* 1615 */                 localImagePlus.unlock();
	/* 1616 */                 localIDataSet.release();
	/* 1617 */                 localIDataSet = null;
	/* 1618 */                 return false;
	/*      */               }
	/*      */ 
	/* 1621 */               IJ.showMessage("Error", "Unknown image type");
	/* 1622 */               localImagePlus.unlock();
	/* 1623 */               localIDataSet.release();
	/* 1624 */               localIDataSet = null;
	/* 1625 */               return false;
	/*      */             }
	/*      */ 
	/* 1628 */             ((ImageProcessor)localObject1).flipVertical();
	/*      */ 
	/* 1630 */             IJ.showProgress(i8, localImageStack.getSize() - 1);
	/*      */           }
	/*      */         }
	/*      */         catch (Exception localException2) {
	/* 1634 */           localException2.printStackTrace();
	/* 1635 */           IJ.showMessage("Error", "Unknown Exception occured while transfer");
	/* 1636 */           localImagePlus.unlock();
	/* 1637 */           localIDataSet.release();
	/* 1638 */           localIDataSet = null;
	/* 1639 */           return false;
	/*      */         }
	/* 1641 */         localImagePlus.unlock();
	/* 1642 */         localIDataSet.release();
	/* 1643 */         localIDataSet = null;
	/* 1644 */         return true;
	/*      */       }
	/* 1646 */       return false;
	/*      */     }	
} //  @jve:decl-index=0:visual-constraint="10,10"
