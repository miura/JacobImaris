import com.jacob.activeX.ActiveXComponent;

import emblcmci.ImarisRemoteExport;

import ij.plugin.PlugIn;


public class Imaris_Jacob implements PlugIn {

	@Override
	//simple test plugin with interface
	public void run(String arg) {
		ImarisRemoteExport view = new ImarisRemoteExport();
		view.setVisible(true);
	}
}

