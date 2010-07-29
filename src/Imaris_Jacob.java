import helloImaris.ImarisRemoteControl;
import ij.plugin.PlugIn;


public class Imaris_Jacob implements PlugIn {

	@Override
	public void run(String arg) {
		ImarisRemoteControl view = new ImarisRemoteControl();
		view.setVisible(true);
	}
}
