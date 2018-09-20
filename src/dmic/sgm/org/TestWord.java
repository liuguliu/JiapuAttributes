package dmic.sgm.org;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hwpf.HWPFDocument;

public class TestWord {
	// ���徲̬����
	static People p[] = new People[500];
	
	public static void main(String[] args) {
		
		String[] commonstring = {"��δ��","����δ��","������ʧ��","������ʧ��","������δ��","��������ʧ��","����������δ��","������ʧ��"};
		
		String text="";
		String filePath = "data";
		String fileName = "12.doc";
        String realPath=filePath+"/"+fileName;	//ƴ��Ϊ�����ֵ�·��
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
                //���word�ĵ����е��ı�  
                //System.out.println(extractor.getText()); 
                text = extractor.getText().toString().trim();
        } catch (Exception e) {
            e.printStackTrace();
        }
        String[] lines = text.split("\r");
        int num = lines.length;
        for(int i=0; i<lines.length; i++) {
        	//System.out.println(lines[i]);
        	
        	p[i] = new People();	//�½�����
        	p[i].id = i+1;
        	p[i].description = lines[i];
        	//��ȡ��ż��Ϣ
        	int indexofpei = lines[i].indexOf("��");
        	int indexofend;
        	if(lines[i].contains("����")) {
        		indexofend = lines[i].indexOf("����");
        	}
        	else if(lines[i].contains("��Ů")) {
        		indexofend = lines[i].indexOf("��Ů")+1;
        	}
        	else if(lines[i].contains("����")) {
        		indexofend = lines[i].indexOf("����")+1;
        	}
        	else {
        		indexofend = lines[i].length();
        	}
        	//System.out.println(indexofpei+" "+indexofend);
        	if(indexofpei!=-1) {
        		p[i].spouseinfo+=lines[i].substring(indexofpei, indexofend);        		
        	}
        	if(lines[i].contains("����")) {
        		int indexbi = lines[i].indexOf("��");
        		p[i].spouseinfo+=lines[i].substring(indexbi+1);
        	}
        	if(p[i].spouseinfo!="") {
        		String[] couple = p[i].spouseinfo.split("��");
            	p[i].spouseinfo="";
            	for(int j=0;j<couple.length;j++) {	//�������������Ϣ
            		if(couple[j].contains("��")&&couple[j].contains("��")) {
            			break;
            		}
            		else {
            			p[i].spouseinfo+=couple[j]+"��";
            		}
            	} 
        	}     	      	
        	
        	
        	String[] lineAttrs = lines[i].split("��");
        	for(int j=0;j<lineAttrs.length;j++) {	//��������������ż��Ϣ����Ů��Ϣ�ֿ�
        		if(lineAttrs[j].charAt(0)=='��') {
        			continue;
        		}
        		else if(p[i].spouseinfo.contains(lineAttrs[j])) {
        			continue;
        		}
        		else if(lineAttrs[j].contains("����")) {	//�ж��Ƿ�Ϊ������Ϣ
        			p[i].soninfo=lineAttrs[j];
        			continue;
        		}
        		else if(WhetherDaughter(lineAttrs[j])) {	//�ж��Ƿ�ΪŮ����Ϣ
        			p[i].daughterinfo=lineAttrs[j];
        			continue;
        		}
        		else {
        			p[i].personInfo+=lineAttrs[j]+"��";
        		}
        	}	//for
        	System.out.println(p[i].personInfo);
        	//System.out.println(p[i].spouseinfo);
        	//System.out.println(p[i].soninfo);
        	//System.out.println(p[i].daughterinfo);
        	//System.out.println("��");
        }	//for
        for(int i=0;i<num;i++) {
        	
         }
	}
	
	public static boolean WhetherDaughter(String lineAttr) {
		if(lineAttr.contains("Ůһ")||lineAttr.contains("Ů��")||lineAttr.contains("Ů��")||lineAttr.contains("Ů��")
				||lineAttr.contains("Ů��")||lineAttr.contains("Ů��")||lineAttr.contains("Ů��")||lineAttr.contains("Ů��")
				||lineAttr.contains("Ů��")||lineAttr.contains("Ůʮ")){
			return true;
		}
		else {
			return false;
		}
	}
	public static void getName(int i, String personinfo) {
		int indexSpace = personinfo.indexOf(' ');	//�����е�������ȷ��
		p[i].name = personinfo.substring(0, indexSpace);
	}
	public static void getFatherName(int i, String personinfo) {
		int indexStart = personinfo.indexOf(' ');
		int indexEnd = personinfo.indexOf("�� ");
		String fartherAndRank = personinfo.substring(indexStart, indexEnd);
		
	}
	public static void getFamilyRank(int i, String fartherAndRank) {
		
	}
}
