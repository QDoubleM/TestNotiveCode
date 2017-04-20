package cn.itcast;

import org.xvolks.jnative.JNative;
import org.xvolks.jnative.exceptions.NativeException;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.sun.jna.Native;
import com.sun.jna.win32.StdCallLibrary;

public class TestNative {

	public native void sayHello();

	/*public interface VISOCX extends StdCallLibrary {
		VISOCX saveas = (VISOCX) Native.loadLibrary("VISOCX",VISOCX.class);

		void saveas();
	}*/

	/*public static void main(String[] args) {
		System.loadLibrary("nativeCode");
		TestNative tst = new TestNative();
		tst.sayHello();
		DrawingControl.saveas.saveas();
	}*/

	public native void DLLGetClassObject();
   // public native void ActivateActCtx();
	public static void main(String[] args) throws NativeException {
		// System.out.println("调用本地API");
	    System.loadLibrary("actxprxy");
	    //System.loadLibrary("Microsoft.Office.Interop.Visio");
	   // VISOCX.saveas.saveas();
		TestNative tst=new TestNative();
		//tst.DLLGetClassObject();

		// System.out.println(System.getProperty("java.library.path"));
		//sJNative visiosaveas = new JNative("nativeCode", "sayHello");
		//System.out.println(((JNative) visiosaveas).isWindows());
		openDoc("F:/张娟的毕设相关/文献摘要3.0-张娟-1210704424.doc");
		saveFileAs("F:/htmlfiles/test1.html");
	}

	 static Dispatch doc = null;

	 static ActiveXComponent ax = new ActiveXComponent("Word.Application");
	 public static void setVisible(boolean visible) {
			ax.setProperty("Visible", new Variant(visible));
		}
	 public static void openDoc(String docPath) {
		    setVisible(false);//是否打开应用
		    Dispatch documents = null;
			documents = Dispatch.get(ax, "Documents").toDispatch();
			doc = Dispatch.call(documents, "open", docPath,new Variant(true),new Variant(false)).toDispatch(); 
			
		}
	 
	 public static void saveFileAs(String filename) {
		 
			Dispatch.call(doc, "SaveAs", filename);
		}
}
