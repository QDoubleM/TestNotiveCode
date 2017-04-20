package cn.itcast;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xdgf.usermodel.XmlVisioDocument;
import org.apache.poi.xdgf.usermodel.shape.ShapeRenderer;
import org.apache.poi.xdgf.util.VsdxToPng;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class VisioFilesToPng implements VisioFileToPng {
	private ActiveXComponent msVisioApp = null;
	private Dispatch document = null;
	public void openVisio(boolean makeVisible) {
		try {
			if (msVisioApp == null) {
				msVisioApp = new ActiveXComponent("Visio.Application");
			}
			Dispatch.put(msVisioApp, "Visible", new Variant(makeVisible));
		} catch (RuntimeException e) {
			e.printStackTrace();
		}
	}
			
	public void openDocument(String visiofilePath) {
		Dispatch documents = Dispatch.get(msVisioApp, "Documents").toDispatch();
		document = Dispatch.call(documents, "Open", visiofilePath).toDispatch();
		}
	
	public void savePageAs(String visioFilePath, String pngFilePath) {

		Dispatch pages = Dispatch.get(document, "Pages").toDispatch();
		// 得到Page的数量
		int pagesCount = Integer.parseInt(Dispatch.get(pages, "Count").toString());
		// 循环得到每个Page
		//String pngpath = new File(visioFilePath).getParent().toString();
		for (int i = 1; i <= pagesCount; i++) {

			Dispatch page = Dispatch.call(pages, "Item", new Variant(i)).toDispatch();
			//getShapes(page);
			String pageName = Dispatch.get(page, "Name").toString();
			Dispatch.call(page, "Export", new Object[] { pngFilePath + "//"+ i + "_" + pageName + ".png" });			 
		}
	}
	public void closeVisio() {
		Dispatch.call(msVisioApp, "Quit");
		msVisioApp = null;
		document = null;
	}
	public static void poiVsdxToPng(String visiofilepath, String pngpath)throws Exception {
		File visiofile = new File(visiofilepath);
		if (visiofile.getName().endsWith("vsdx")) {
			XmlVisioDocument visioDocument = new XmlVisioDocument(new FileInputStream(visiofilepath));
			ShapeRenderer renderer = new ShapeRenderer();
			VsdxToPng.renderToPng(visioDocument, pngpath, 181.81818181818181D,renderer);// 181.***这差不多是缩放比
		} else
			System.out.println("文件类型不支持，需要文件类型为vsdx");
	}	
	@Override
	public void VsdxFileToPng(String visiofilepath, String pngpath)throws Exception {		
        String ostype = System.getProperty("os.name");		
		if (ostype.toLowerCase().startsWith("windows")) {
			System.out.println("当前操作系统为windows");
			openVisio(false); // 设定Visio Application是否打开
			openDocument(visiofilepath); // 建立文件内容			
			savePageAs(visiofilepath, pngpath);
			closeVisio();
			System.out.println("在windows下利用jacob，visio文件转换png图片完成");
		} else if (ostype.toLowerCase().startsWith("linux")) {
			System.out.println("当前操作系统为Linux");
			poiVsdxToPng(visiofilepath, pngpath);
			System.out.println("在linux下利用jacob，visio文件转换png图片完成");
		}
	}
}
