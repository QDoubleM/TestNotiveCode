package cn.itcast;

public interface WordsToHtml {

	public void poiWordToHtml(String wordfilepath,String htmlfilepath) throws Exception;
	//wordfilepath是文档的完整路径
	//htmlfilepath是生成的html文件存放的目标文件夹
	public void jacobWordToHtml(String wordfilepath,String htmlfilepath) throws Exception;
}
