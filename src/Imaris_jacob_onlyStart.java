import com.jacob.activeX.ActiveXComponent;

import emblcmci.GrayHyperStackSplitter;

import ij.ImagePlus;
import ij.WindowManager;
import ij.plugin.PlugIn;


public class Imaris_jacob_onlyStart implements PlugIn {
	
	private ActiveXComponent imarisApplication;
	
	public void run(String arg) {
		
//		if (imarisApplication == null) {
//			imarisApplication = new ActiveXComponent(
//					"Imaris.Application");								
//			imarisApplication.setProperty("mVisible", true);
//			imarisApplication.setProperty("mUserControl", true); //this keeps Imaris Opened
//		}
		ImagePlus imp = WindowManager.getCurrentImage();
		GrayHyperStackSplitter splitter = new GrayHyperStackSplitter(imp);
		ImagePlus imp0 = splitter.extractZTfromCZT(imp, 0);
		ImagePlus imp1 = splitter.extractZTfromCZT(imp, 1);
		imp0.show();
		imp1.show();
	}
	
}
