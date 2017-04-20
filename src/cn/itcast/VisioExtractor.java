package cn.itcast;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class VisioExtractor {
	// 创建一个组件。
		private ActiveXComponent msVisioApp = null;

		// 整个模板
		private Dispatch document = null;

		// 选中的控件
		private Dispatch session = null;

		// 构造函数
		VisioExtractor() {
			super();
		}

		/**
		 * 开启visio档案
		 * 
		 * @param makeVisible
		 *            显示或是不显示(true:显示;false:不显示)
		 */
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

		/**
		 * 建立viso的文本内容
		 * 
		 */
		public void createNewDocument() {
			// 得到
			Dispatch documents = msVisioApp.getProperty("Documents").toDispatch();
			// 添加 document world 可以 ， visio 不可以 可能是缺少参数
			Dispatch document = Dispatch.call(documents, "add").toDispatch();
			// 添加 page页签
			Dispatch page = Dispatch.call(document, "add").toDispatch();
			// 添加 页签选中 当前页签
			Dispatch.call(page, "Select");

		}

		public void openDocument(String _filePath) {

			Dispatch documents = Dispatch.get(msVisioApp, "Documents").toDispatch();
			// 打开viso文件
			document = Dispatch.call(documents, "Open", _filePath).toDispatch();
			// 得到所有的Pages
			Dispatch pages = Dispatch.get(document, "Pages").toDispatch();
		}

		/**
		 * 添加一个页签
		 */
		public void addPage() {

			Dispatch pages = Dispatch.get(document, "Pages").toDispatch();

			session = Dispatch.call(pages, "add").toDispatch();

			// Dispatch.call(session, "Name", "XXX流程"); // 写入标题内容 // 标题格行
		}

		/**
		 * 获得页签集合 说明：
		 * 
		 * @return
		 * @throws Exception
		 *             创建时间：2011-6-4 下午05:33:07
		 */
		public void getShapes(Dispatch page) {
			// Shapes/Shape
			Dispatch vshapes = Dispatch.get(page, "Shapes").toDispatch();
			// 得到Shapes的数量
			int pagesCount = Integer.parseInt(Dispatch.get(vshapes, "Count")
					.toString());
			for (int i = 1; i <= pagesCount; i++) {
				Dispatch Shape = Dispatch.call(vshapes, "Item", new Variant(i))
						.toDispatch();

				String shapeid = Dispatch.get(Shape, "Id").toString();
				String shapetype = Dispatch.get(Shape, "Type").toString();
				String shapetext = Dispatch.get(Shape, "Text").toString();
				// String shapename = Dispatch.get(Shape, "Name").toString() ;
				// String shapenameU = Dispatch.get(Shape, "NameU").toString() ;
				// String shapelineStyle = Dispatch.get(Shape,
				// "LineStyle").toString() ;
				// String shapefillStyle = Dispatch.get(Shape,
				// "FillStyle").toString() ;
				// String shapetextStyle = Dispatch.get(Shape,
				// "TextStyle").toString() ;
				System.out.print("    " + i + "shape id:" + shapeid);
				System.out.print("    " + i + "shape type:" + shapetype);
				System.out.print("    " + i + "shape text:" + shapetext);
				System.out.println();
			}

		}

		public void documentToString() {
			Dispatch pages = Dispatch.get(document, "Pages").toDispatch();
			// 得到Page的数量
			int pagesCount = Integer.parseInt(Dispatch.get(pages, "Count")
					.toString());

			System.out.println("图片数量：" + pagesCount);
			// 循环得到每个Page
			for (int i = 1; i <= pagesCount; i++) {

				Dispatch page = Dispatch.call(pages, "Item", new Variant(i))
						.toDispatch();

				String pageid = Dispatch.get(page, "Id").toString();
				String pagename = Dispatch.get(page, "Name").toString();
				String pagenameU = Dispatch.get(page, "NameU").toString();

				System.out.print(i + " page id:" + pageid);
				System.out.print(i + " page name:" + pagename);
				System.out.print(i + " page nameU:" + pagenameU);

				getShapes(page);
			}
		}

		/**
		 * 另存为
		 * 
		 * @param type
		 */
		public void savePageAs(String visioFilePath, String type) {

			Dispatch pages = Dispatch.get(document, "Pages").toDispatch();
			// 得到Page的数量
			int pagesCount = Integer.parseInt(Dispatch.get(pages, "Count")
					.toString());

			System.out.println("图片数量：" + pagesCount);
			String pngpath = new File(visioFilePath).getParent().toString();
			// 循环得到每个Page
			for (int i = 1; i <= pagesCount; i++) {

				Dispatch page = Dispatch.call(pages, "Item", new Variant(i))
						.toDispatch();
				getShapes(page);
				// 输出Page的名称
				String pageName = Dispatch.get(page, "Name").toString();

				if ("png".equals(type)) {
					// 将该Page保存为图片
					Dispatch.call(page, "Export", new Object[] { pngpath + "//"
							+ i + "_" + pageName + ".png" });
				}

			}

		}

		/**
		 * 关闭文本内容(如果未开启visio编辑时,释放ActiveX执行绪)
		 */
		public void closeDocument() {
			// visio 的关闭， 没有参数或者参数不对。 一致没有找到不保存关闭的方法。
			Dispatch.call(document, "Save");
			Dispatch.call(document, "Close");
			document = null;
		}

		/**
		 * 关闭visio(如果未开启visio编辑时,释放ActiveX执行绪)
		 */
		public void closeVisio() {
			Dispatch.call(msVisioApp, "Quit");
			msVisioApp = null;
			document = null;
		}

		/**
		 * @param args
		 */
		public static void main(String[] args) {

			String otFile = "F:/visiofiles/test1.vsdx";

			VisioExtractor visio = new VisioExtractor(); // 建立一个VisioExtractor对象
			visio.openVisio(true); // 设定Visio开启显示
			// visio.createNewDocument();
			visio.openDocument(otFile); // 建立文件内容
			visio.addPage();
			visio.documentToString();
			// visio.closeDocument();
			// visio.closeVisio();
			visio.savePageAs(otFile, "png");

		}
}
