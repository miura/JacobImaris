import com.jacob.activeX.ActiveXComponent;

import helloImaris.ImarisRemoteControl;
import ij.plugin.PlugIn;


public class Imaris_Jacob implements PlugIn {

	@Override
	//simple test plugin with interface
	public void run(String arg) {
		ImarisRemoteControl view = new ImarisRemoteControl();
		view.setVisible(true);
	}
}

