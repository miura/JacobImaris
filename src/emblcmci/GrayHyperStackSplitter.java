package emblcmci;

import ij.ImagePlus;
import ij.ImageStack;

public class GrayHyperStackSplitter {
	ImagePlus imp;
	ImagePlus impsplit;

	/**
	 * @param imp
	 */
	public GrayHyperStackSplitter(ImagePlus imp) {
		super();
		this.imp = imp;
	}
	
	/**
	 * 
	 * @param imp
	 * @param Ch: starts from 0
	 */
	public void splitter(ImagePlus imp, int Ch){
		this.imp = imp;
		int nFrames = imp.getNFrames();
		int nSlices = imp.getNSlices();
		int nChannels = imp.getNChannels();
		
		if (!imp.isHyperStack()) return;
		if (Ch > (nChannels - 1)) return;
	
		// following is a modification of HyperstackConverter.java

		ImageStack stack = imp.getImageStack();
		
		ImageStack splitstack = new ImageStack(stack.getWidth(), stack.getHeight(), stack.getSize()/imp.getNChannels());
		
		Object[] images1 = stack.getImageArray();
		Object[] images3 = splitstack.getImageArray();
		
		String[] labels1 = stack.getSliceLabels();
		String[] labels3 = splitstack.getSliceLabels();
		int[] index = new int[3];
		for (index[2]=0; index[2]<nFrames; ++index[2]) {
			for (index[1]=0; index[1]<nSlices; ++index[1]) {
				//for (index[0]=0; index[0]<nChannels; ++index[0]) {
					index[0] = Ch;
					int dstIndex = 0 + index[1] + index[2] * nSlices;
					int srcIndex = index[0] + index[1] * nChannels + index[2]* nChannels * nSlices;
					images3[dstIndex] = images1[srcIndex];
					labels3[dstIndex] = labels1[srcIndex];
				//}
			}
		}
		impsplit = new ImagePlus("splitted", splitstack);
		impsplit.setOpenAsHyperStack(true);
		impsplit.setDimensions(1, nSlices, nFrames);
		impsplit.setCalibration(imp.getCalibration());
	}
	
	public ImagePlus extractZTfromCZT(ImagePlus imp, int Ch){
		splitter(imp, Ch);
		return impsplit;
	}
	
	
}
