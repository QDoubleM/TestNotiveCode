package cn.itcast;

import java.io.File;

import jp.ne.so_net.ga2.no_ji.jcom.IDispatch;
import jp.ne.so_net.ga2.no_ji.jcom.JComException;
import jp.ne.so_net.ga2.no_ji.jcom.ReleaseManager;

public class ExportToPicsFromVisio {
	private void createDir(String outPath){  
        File file = new File(outPath);  
        if(file.exists()){  
            file.delete();  
        }  
        try {  
            file.mkdir();  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    }  
    private void visioTest(String vsdFilePath, String outPath) {  
        createDir(outPath);  
        ReleaseManager rm = new ReleaseManager();  
        IDispatch visioApp;  
        try {  
            // 调用Visio程序  
            visioApp = new IDispatch(rm, "Visio.Application");  
            // 为了方便程序调试，设置成了显示打开Visio，正式用改成false  
            visioApp.put("Visible", new Boolean(false));  
            IDispatch documents = (IDispatch) visioApp.get("Documents");  
            // 打开文件  
            IDispatch doc = (IDispatch) documents.method("open", new Object[] { vsdFilePath });  
            // 得到所有的Pages  
            IDispatch pages = (IDispatch) doc.get("Pages");  
            // 得到Page的数量  
            int pagesCount = Integer.parseInt(pages.get("Count").toString());  
            System.out.println("图片数量："+pagesCount);  
            // 循环得到每个Page  
            for (int i = 1; i <= pagesCount; i++) {  
                IDispatch page = (IDispatch) pages.method("item",  
                        new Object[] { new Integer(i) });  
                // 输出Page的名称  
                System.out.println(page.get("Name"));  
                // 将该Page保存为图片  
                page.method("Export", new Object[] { outPath + i+"_"+page.get("Name") + ".jpg" });  
            }  
            //Thread.sleep(5000);  
            // Quit without saving  
            visioApp.method("quit", null);  
            visioApp.release();  
  
        } catch (JComException e) {  
            // TODO Auto-generated catch block  
            e.printStackTrace();  
        /*} catch (InterruptedException e) { 
            // TODO Auto-generated catch block 
            e.printStackTrace();*/  
        }  
  
    }  
    public static void main(String[] args) {  
        // TODO Auto-generated method stub  
        ExportToPicsFromVisio v = new ExportToPicsFromVisio();  
        v.visioTest("F:\\visiofiles\\test1.vsdx", "F:\\pngs\\");  
  
    }  
}
