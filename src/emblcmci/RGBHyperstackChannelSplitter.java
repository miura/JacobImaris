package emblcmci;
/* 20100804 
 * Kota Miura (miura@embl.de)
 * Split RGB-Z or RGB-T or RGB-ZT stack.
 *  
 * coded for exporting such Hyperstacks to Imaris. 
 */
import ij.IJ;
import ij.ImagePlus;
import ij.measure.Calibration;
import ij.plugin.filter.RGBStackSplitter;

public class RGBHyperstackChannelSplitter {
	private ImagePlus imp;
	public ImagePlus impCh1;
	public ImagePlus impCh2;
	public ImagePlus impCh3;	
	public int chnum = 3;
	int srcFrames = 0;
	int srcSlices = 0;
	int srcChannels = 0;
	Calibration cal = null;
	
	/**
	 * @param imp
	 */
	public RGBHyperstackChannelSplitter(ImagePlus imp) {
		super();
		this.imp = imp;
		getDimensions();
		cal = new Calibration(imp);
	}
	
	void getDimensions(){
		srcFrames = imp.getNFrames();
		srcChannels = imp.getNChannels();
		srcSlices = imp.getNSlices();		
	}
	
	public boolean ChannelSplitter(ImagePlus imp){
		this.imp = imp;
		getDimensions();
		cal = new Calibration(imp);
		if (imp.getStackSize() == 1) {
			IJ.error("Splitting cannot be done with single image");
			return false;
		}
		RGBStackSplitter splitter = new RGBStackSplitter();
		if (imp.getBitDepth() == 24) {
			splitter.split(imp);
			impCh1 = new ImagePlus("Ch1", splitter.red);
			impCh2 = new ImagePlus("Ch2", splitter.green);
			impCh3 = new ImagePlus("Ch3", splitter.blue);		

			impCh1.setDimensions(1, srcSlices, srcFrames);
			impCh2.setDimensions(1, srcSlices, srcFrames);
			impCh3.setDimensions(1, srcSlices, srcFrames);

			impCh1.setCalibration(cal);
			impCh2.setCalibration(cal);
			impCh3.setCalibration(cal);
			return true;
		} else return false;
		
	}
	
	public ImagePlus getChannel1(){
		return impCh1;
	}
	public ImagePlus getChannel2(){
		return impCh2;
	}	
	public ImagePlus getChannel3(){
		return impCh2;
	}		
}
