package dmic.sgm.org;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hwpf.HWPFDocument;

public class TestWord {
	// 定义静态变量
	static People p[] = new People[500];
	
	public static void main(String[] args) {
		
		String[] commonstring = {"生未详","生卒未详","公卒葬失考","妣卒葬失考","妣卒葬未详","公妣卒葬失考","公与王妣卒未详","生卒葬失考"};
		
		String text="";
		String filePath = "data";
		String fileName = "12.doc";
        String realPath=filePath+"/"+fileName;	//拼接为含名字的路径
        /*File file = new File(realPath);
        if(file.exists()) {
        	System.out.println("YES");
        }
        System.out.println(realPath);*/
        try {
                //FileInputStream in= new FileInputStream(realPath);
                //WordExtractor extractor = new WordExtractor(in);
                InputStream is = new FileInputStream(realPath);  
                //WordExtractor extractor = new WordExtractor(is);  
                HWPFDocument extractor = new HWPFDocument(is); 
                //输出word文档所有的文本  
                //System.out.println(extractor.getText()); 
                text = extractor.getText().toString().trim();
        } catch (Exception e) {
            e.printStackTrace();
        }
        String[] lines = text.split("\r");
        int num = lines.length;
        for(int i=0; i<lines.length; i++) {
        	//System.out.println(lines[i]);
        	
        	p[i] = new People();	//新建人物
        	p[i].id = i+1;
        	p[i].description = lines[i];
        	//提取配偶信息
        	int indexofpei = lines[i].indexOf("配");
        	int indexofend;
        	if(lines[i].contains("生子")) {
        		indexofend = lines[i].indexOf("生子");
        	}
        	else if(lines[i].contains("。女")) {
        		indexofend = lines[i].indexOf("。女")+1;
        	}
        	else if(lines[i].contains("。公")) {
        		indexofend = lines[i].indexOf("。公")+1;
        	}
        	else {
        		indexofend = lines[i].length();
        	}
        	//System.out.println(indexofpei+" "+indexofend);
        	if(indexofpei!=-1) {
        		p[i].spouseinfo+=lines[i].substring(indexofpei, indexofend);        		
        	}
        	if(lines[i].contains("。妣")) {
        		int indexbi = lines[i].indexOf("妣");
        		p[i].spouseinfo+=lines[i].substring(indexbi+1);
        	}
        	if(p[i].spouseinfo!="") {
        		String[] couple = p[i].spouseinfo.split("。");
            	p[i].spouseinfo="";
            	for(int j=0;j<couple.length;j++) {	//清除公妣共有信息
            		if(couple[j].contains("公")&&couple[j].contains("妣")) {
            			break;
            		}
            		else {
            			p[i].spouseinfo+=couple[j]+"。";
            		}
            	} 
        	}     	      	
        	
        	
        	String[] lineAttrs = lines[i].split("。");
        	for(int j=0;j<lineAttrs.length;j++) {	//将人物属性与配偶信息及子女信息分开
        		if(lineAttrs[j].charAt(0)=='妣') {
        			continue;
        		}
        		else if(p[i].spouseinfo.contains(lineAttrs[j])) {
        			continue;
        		}
        		else if(lineAttrs[j].contains("生子")) {	//判断是否为儿子信息
        			p[i].soninfo=lineAttrs[j];
        			continue;
        		}
        		else if(WhetherDaughter(lineAttrs[j])) {	//判断是否为女儿信息
        			p[i].daughterinfo=lineAttrs[j];
        			continue;
        		}
        		else {
        			p[i].personInfo+=lineAttrs[j]+"。";
        		}
        	}	//for
        	System.out.println(p[i].personInfo);
        	//System.out.println(p[i].spouseinfo);
        	//System.out.println(p[i].soninfo);
        	//System.out.println(p[i].daughterinfo);
        	//System.out.println("完");
        }	//for
        for(int i=0;i<num;i++) {
        	
         }
	}
	
	public static boolean WhetherDaughter(String lineAttr) {
		if(lineAttr.contains("女一")||lineAttr.contains("女二")||lineAttr.contains("女三")||lineAttr.contains("女四")
				||lineAttr.contains("女五")||lineAttr.contains("女六")||lineAttr.contains("女七")||lineAttr.contains("女八")
				||lineAttr.contains("女九")||lineAttr.contains("女十")){
			return true;
		}
		else {
			return false;
		}
	}
	public static void getName(int i, String personinfo) {
		int indexSpace = personinfo.indexOf(' ');	//名字中的字数不确定
		p[i].name = personinfo.substring(0, indexSpace);
	}
	public static void getFatherName(int i, String personinfo) {
		int indexStart = personinfo.indexOf(' ');
		int indexEnd = personinfo.indexOf("子 ");
		String fartherAndRank = personinfo.substring(indexStart, indexEnd);
		
	}
	public static void getFamilyRank(int i, String fartherAndRank) {
		
	}
}
