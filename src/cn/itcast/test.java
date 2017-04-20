package cn.itcast;

public class test { 
	public static void main(String args[]) throws Exception{
		//String wordfilepath = "F:/毕业实习相关/概率统计随机过程复习.docx";
		//String htmlfilepath = "F:/htmlfiles";	
		WordsToHtml wth = null;
		WordFileToHtml pwth = new WordFileToHtml();
		wth = pwth;
		//wth.WordToHtml("F:/毕业实习相关/概率统计随机过程复习.docx", "F:/htmlfiles");
        wth.jacobWordToHtml("F:/毕业实习相关/概率统计随机过程复习.docx", "F:/htmlfiles");
		VisioFileToPng visiotopng = null;
		VisioFilesToPng visiofilestopng = new VisioFilesToPng();
		visiotopng = visiofilestopng;
		visiotopng.VsdxFileToPng("F:\\visiofiles\\test1.vsdx", "F:\\pngs");
	}
}
