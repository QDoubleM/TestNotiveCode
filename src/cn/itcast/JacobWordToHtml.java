package cn.itcast;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public abstract class JacobWordToHtml implements WordsToHtml {
	 public static boolean wordToHtml (String wordfilePath,String htmlfilePath) {  
		  String filename = new File(wordfilePath).getName().toString();
		  //启动word  
		 ActiveXComponent activexcomponent = new ActiveXComponent("Word.Application");  		    
		  boolean flag = false;  		    
		  try {  
		   //设置word不可见  
		   activexcomponent.setProperty("Visible",new Variant(false));  		     
		   Dispatch docs = activexcomponent.getProperty("Documents").toDispatch();  		     
		   //打开word文档  
		   Dispatch doc = Dispatch.invoke(docs,"Open",Dispatch.Method,new Object[]{wordfilePath,new Variant(false), new Variant(true)},  
		     new int[1]).toDispatch();  
		     
		   Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { htmlfilePath+"/"+filename.substring(0, filename.lastIndexOf("."))+".html", new Variant(8) }, new int[1]);
		   Variant f = new Variant(false);  
		   Dispatch.call(doc, "Close", f);  
		   flag = true;  
		   return flag;  
		     
		  } catch (Exception e) {  
		   e.printStackTrace();  
		   return flag;  
		  } finally{  
			  activexcomponent.invoke("Quit", new Variant[] {});  
		  }  
		 }  
		   
		
		@Override
		public void jacobWordToHtml(String wordfilepath, String htmlfilepath)throws Exception {
			JacobWordToHtml.wordToHtml(wordfilepath, htmlfilepath);  
			 System.out.println("转换完成");
		}  
}
