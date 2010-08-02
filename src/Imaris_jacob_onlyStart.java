import com.jacob.activeX.ActiveXComponent;
import ij.plugin.PlugIn;


public class Imaris_jacob_onlyStart implements PlugIn {
	
	private ActiveXComponent imarisApplication;
	
	public void run(String arg) {
		
		if (imarisApplication == null) {
			imarisApplication = new ActiveXComponent(
					"Imaris.Application");								
			imarisApplication.setProperty("mVisible", true);
			imarisApplication.setProperty("mUserControl", true); //this keeps Imaris Opened
		}
	}
	
}
