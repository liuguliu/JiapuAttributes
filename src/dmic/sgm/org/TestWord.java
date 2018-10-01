package dmic.sgm.org;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hwpf.HWPFDocument;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class TestWord {
	// ���徲̬����
	static People p[] = new People[1000];
	static String cNum[] = {"һ","��","��","��","��","��","��","��","��","ʮ"};
	static String rankNum[] = {"","��","��","��","��","��","��","��","��","��","ʮ"};
	static String generationRank[] = {"һ","��","��","��","��","��","��","��","��","ʮ","ʮһ","ʮ��","��","��","��","��","��","ʢ","ʱ"};
	public static void main(String[] args) {
		
		//String[] commonstring = {"��δ��","����δ��","������ʧ��","������ʧ��","������δ��","��������ʧ��","����������δ��","������ʧ��"};
		
		String text="";
		String filePath = "data";
		String fileName = "12.doc";
        String realPath=filePath+"/"+fileName;	//ƴ��Ϊ�����ֵ�·��
        try {
        	if(fileName.endsWith(".doc")) {
        		InputStream is = new FileInputStream(realPath);    
                HWPFDocument extractor = new HWPFDocument(is); 
                //���word�ĵ����е��ı�  
                //System.out.println(extractor.getText()); 
                text = extractor.getText().toString().trim();
        	}
            if(fileName.endsWith(".txt")) {
            	FileReader filereader = new FileReader(realPath);
        		BufferedReader bufferedreader = new BufferedReader(filereader);
        		String line="";
        		while((line=bufferedreader.readLine())!=null){
                    text+=line+"\r";
                }
        		System.out.println(text);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        String[] lines = text.split("\r");
        int num = lines.length;
        System.out.println(num);
        for(int i=0; i<lines.length; i++) {	//1 ѭ����ȡÿһ�е�����
        	//System.out.println(lines[i]);
        	p[i] = new People();	//�½�����
        	p[i].id = i+1; 
        	p[i].description = lines[i];
//        	//2 ��ȡ��ż��Ϣ
//        	int indexofpei;
//        	if(lines[i].contains("��")) {
//        		indexofpei=lines[i].indexOf("��");
//        	}else if(lines[i].contains("��")){
//        		indexofpei=lines[i].indexOf("��");
//        	}else {
//        		indexofpei = -1;
//        	}
//        	int indexofend;
//        	if(lines[i].contains(" ��")) {
//        		indexofend = lines[i].indexOf(" ��");
//        	}
//        	else if(lines[i].contains(" Ů")) {
//        		indexofend = lines[i].indexOf(" Ů")+1;
//        	}
//        	else if(lines[i].contains(" ��")) {
//        		indexofend = lines[i].indexOf(" ��")+1;
//        	}
//        	else {
//        		indexofend = lines[i].length();
//        	}
//        	//System.out.println(indexofpei+" "+indexofend);
//        	if(indexofpei!=-1) {
//        		p[i].spouseinfo+=lines[i].substring(indexofpei, indexofend);        		
//        	}
//        	if(lines[i].contains("����")) {
//        		int indexbi = lines[i].indexOf("��");
//        		p[i].spouseinfo+=lines[i].substring(indexbi+1);
//        	}
//        	if(p[i].spouseinfo!="") {
//        		String[] couple = p[i].spouseinfo.split("��");
//            	p[i].spouseinfo="";
//            	for(int j=0;j<couple.length;j++) {	//�������������Ϣ
//            		if(couple[j].contains("��")&&couple[j].contains("��")) {
//            			break;
//            		}
//            		else {
//            			p[i].spouseinfo+=couple[j]+"��";
//            		}
//            	} 
//        	}     	      	
        	
        	String[] lineAttrs = lines[i].split("��");
        	for(int j=0;j<lineAttrs.length;j++) {	//3 ��������������ż��Ϣ����Ů��Ϣ�ֿ�
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
        	//System.out.println(p[i].personInfo);
        	//System.out.println(p[i].spouseinfo);
        	//System.out.println(p[i].soninfo);
        	//System.out.println(p[i].daughterinfo);
        	//System.out.println("��");
        	//4 ��ȡ��������
        	getName(i, p[i].personInfo);
        	getFatherNameAndRank(i, p[i].personInfo);
        	getCourtesyName(i, p[i].personInfo);
        	getpesudonym(i, p[i].personInfo);
        	getBirthday(i, p[i].personInfo);
        	getDeathdate(i, p[i].personInfo);
        	getburied(i, p[i].personInfo);
        	getRank(i);
        	//System.out.println(p[i].toStringTest());
        	//5 ������ż��Ϣ
        	if(p[i].spouseinfo!="") {
        		if(p[i].spouseinfo.contains("����")) {
        			String wife2 = p[i].spouseinfo.substring(p[i].spouseinfo.indexOf("����")+2);
            		String wife1 = p[i].spouseinfo.substring(1, p[i].spouseinfo.indexOf("����"));
            		getWifeINfo(i, num, wife1, "");
            		num++;           		
            		if(p[i].spouseinfo.contains("����")) {
            			 String wife3 = wife2.substring(wife2.indexOf("����")+2);
            			 wife2 = wife2.substring(0,wife2.indexOf("����"));
            			 getWifeINfo(i, num, wife2, "����");
                 		 num++;
                 		 getWifeINfo(i, num, wife3, "����");
                		 num++;
            		}else {
            			 getWifeINfo(i, num, wife2, "����");
                		 num++;
            		}
            	}else {
            		getWifeINfo(i, num, p[i].spouseinfo.substring(1), "");
            		num++;
            	}
        	}
        	getFatheridAndMotherid(i);
        }	//for
        for(int i=0; i<lines.length; i++) {	//1 ѭ����ȡÿһ�е�����
        	//6 ���������Ϣ
        	if(p[i].soninfo!="") {
        		//System.out.println(p[i].soninfo);
        		String[] sons = p[i].soninfo.split(" ");
        		for(int j =1; j<sons.length;j++) {
        			String nameson="";
        			String description = "";
        			//��ȡ����������
        			if(sons[j].contains("(")) {
        				nameson=sons[j].substring(0, sons[j].indexOf("("));
        				description = sons[j].substring(sons[j].indexOf("(")+1,sons[j].length()-1);
        				//System.out.print(nameson+" ");
        			}else{
        				nameson = sons[j];
        				//System.out.print(nameson+" ");
        			}
        			//�ж��Ƿ����
        			boolean flag=false;
        			for(int k=0;k<lines.length;k++) {
        				if((p[k].name.equals(nameson)||p[k].name.equals(generationRank[p[i].generition]+nameson))
        						&&p[k].fathername.equals(p[i].name)) {
        					flag=true;
        					break;
        				}
        			}
        			if(!flag) {	//��������
        				p[num] = new People();
        				p[num].id = num+1;
        				if(nameson.length()==1) {
        					p[num].name = generationRank[p[i].generition]+nameson;
        				}else {
        					p[num].name = nameson;
        				}
        				p[num].generition=p[i].generition+1;
        				p[num].familyrank = j;
        				p[num].fatherid = p[i].id;
        				
        				int motherid = 0;
        				if(p[i].partnerid!="") {
        					if(p[i].partnerid.contains("/")) {
            					motherid=Integer.parseInt(p[i].partnerid.split("/")[0]);
            				}else {
            					motherid=Integer.parseInt(p[i].partnerid);
            				}
            				p[num].motherid = motherid;
        				}     				
        				p[num].description = description;
        				//System.out.println(p[num].toString());
        				num++;
        			}else {
						continue;
					}
        		}
        		//System.out.println();
        	}
        }
        for(int i=0;i<num;i++) {
        	if(p[i].gender.equals("��")) {
        		p[i].name = "��" + p[i].name;
        	}
        }
        for(int i=0;i<num;i++) {    	
        	//7 ����Ů����Ϣ
        	if(p[i].daughterinfo!="") {
        		System.out.println(p[i].daughterinfo);
        		int daughternum = getDaughterNum(p[i].daughterinfo);
        		if(!p[i].daughterinfo.contains(" ")) {   //��ز     			        			
    				for(int j=0;j<daughternum;j++) {
    					p[num] = new People();
    					p[num].id = num+1;
    					p[num].name = "Ů"+cNum[j];
    					p[num].gender="Ů";
    					p[num].familyrank=j+1;
    					p[num].generition=p[i].generition+1;
    					p[num].fatherid = p[i].id;
        				
        				int motherid = 0;
        				if(p[i].partnerid!="") {
        					if(p[i].partnerid.contains("/")) {
            					motherid=Integer.parseInt(p[i].partnerid.split("/")[0]);
            				}else {
            					motherid=Integer.parseInt(p[i].partnerid);
            				}
        				}
        				p[num].motherid = motherid;
        				if(p[i].daughterinfo.contains(")")){
        					p[num].description = "ز";
        				}
        				//System.out.println(p[num].toString());
        				num++;
    				}
        		}else {
        			String[] daughters = p[i].daughterinfo.split(" ");
        			for(int j=1;j<daughternum;j++) {
        				if(daughters[j].contains(rankNum[j]+"��")||daughters[j].charAt(0)=='��') {
        					p[num]=new People();
        					p[num].id=num+1;
        					p[num].name="Ů"+cNum[j-1];
        					p[num].gender="Ů";
        					p[num].familyrank=j;
        					p[num].generition=p[i].generition+1;
        					p[num].fatherid = p[i].id;
            				
            				int motherid = 0;
            				if(p[i].partnerid!="") {
            					if(p[i].partnerid.contains("/")) {
                					motherid=Integer.parseInt(p[i].partnerid.split("/")[0]);
                				}else {
                					motherid=Integer.parseInt(p[i].partnerid);
                				}
                				p[num].motherid = motherid;
            				}            				
            				if(daughters[j].contains(rankNum[j]+"��")) {
            					p[num].description = daughters[j].substring(1);
            				}else {
            					p[num].description = daughters[j];
            				}
            				p[num].partnerid = String.valueOf(num+2);
            				//System.out.println(p[num].toString());
            				num++;
            				//��
            				p[num]=new People();
        					p[num].id=num+1;
        					if(daughters[j].contains(rankNum[j]+"��")) {
        						p[num].name=daughters[j].substring(2);
            				}else {
            					p[num].name=daughters[j].substring(1);
            				}
        					p[num].partnerid = String.valueOf(num);
        					num++;
        				}else  {
        					p[num]=new People();
        					p[num].id=num+1;
        					p[num].name="Ů"+cNum[j-1];
        					p[num].gender="Ů";
        					p[num].familyrank=j;
        					p[num].generition=p[i].generition+1;
        					p[num].fatherid = p[i].id;
            				
            				int motherid = 0;
            				if(p[i].partnerid!="") {
            					if(p[i].partnerid.contains("/")) {
                					motherid=Integer.parseInt(p[i].partnerid.split("/")[0]);
                				}else {
                					motherid=Integer.parseInt(p[i].partnerid);
                				}
                				p[num].motherid = motherid;
            				}
            				if(daughters[j].contains("(�|)")||daughters[j].contains("(ز)")) {
            					p[num].description = "ز";
            				}
            				//System.out.println(p[num].toString());
            				num++;
        				}
        			}
        		}
        	}//7 ����Ů����Ϣ
        }
        String path = "data/��15��.xls";
        List<People> list = new ArrayList<People>();
//        for(int i=0;i<num;i++) {
//        	list.add(p[i]);
//        	System.out.println(p[i].toString());
//        }
        //addExcel(path,list);
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
	public static void getFatherNameAndRank(int i, String personinfo) {
		int indexStart = personinfo.indexOf(' ');
		int indexEnd = personinfo.indexOf("�� ");
		
		String fatherAndRank = personinfo.substring(indexStart, indexEnd+1);
		//System.out.println(indexStart+" "+indexEnd+" "+fatherAndRank);
		int indexofRank = getFamilyRank(i, fatherAndRank);
		//System.out.println(indexofRank);
		p[i].fathername = fatherAndRank.substring(1,indexofRank);

	}
	public static int getFamilyRank(int i, String fartherAndRank) {
		int indexofRank = -1;
		
		if(fartherAndRank.contains("֮��")||fartherAndRank.contains("����")) {
			p[i].familyrank=1;
			if(fartherAndRank.indexOf("֮��")!=-1) {
				indexofRank = fartherAndRank.indexOf("֮��");
			}
			else {
				indexofRank = fartherAndRank.indexOf("����");
			}
		}
		else if(fartherAndRank.contains("����")) {
			p[i].familyrank=2;
			indexofRank = fartherAndRank.indexOf("����");
		}
		else if(fartherAndRank.contains("����")) {
			p[i].familyrank=3;
			indexofRank = fartherAndRank.indexOf("����");
		}
		else if(fartherAndRank.contains("����")) {
			p[i].familyrank=4;
			indexofRank = fartherAndRank.indexOf("����");
		}
		else if(fartherAndRank.contains("����")) {
			p[i].familyrank=5;
			indexofRank = fartherAndRank.indexOf("����");
		}
		else if(fartherAndRank.contains("����")) {
			p[i].familyrank=6;
			indexofRank = fartherAndRank.indexOf("����");
		}
		else if(fartherAndRank.contains("����")) {
			p[i].familyrank=7;
			indexofRank = fartherAndRank.indexOf("����");
		}
		else if(fartherAndRank.contains("����")) {
			p[i].familyrank=8;
			indexofRank = fartherAndRank.indexOf("����");
		}
		else if(fartherAndRank.contains("����")) {
			p[i].familyrank=9;
			indexofRank = fartherAndRank.indexOf("����");
		}
		else if(fartherAndRank.contains("ʮ��")) {
			p[i].familyrank=10;
			indexofRank = fartherAndRank.indexOf("ʮ��");
		}
		return indexofRank;
	}
	public static void getCourtesyName(int i, String personinfo) {
		String pattern = "��.*? ";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(personinfo);
		if(m.find()) {
			String courtesyinfo =m.group(0).trim();
			if(courtesyinfo.contains("����")) {
				int indexofYou = courtesyinfo.indexOf("����");
				p[i].CourtesyName = courtesyinfo.substring(1, indexofYou);
				p[i].CourtesyName = p[i].CourtesyName + "/"+courtesyinfo.substring(indexofYou+2);
			}
			else {
				p[i].CourtesyName = courtesyinfo.substring(1);
			}
		}
	}
	public static void getpesudonym(int i, String personinfo) {
		String pattern = "��.*? ";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(personinfo);
		if(m.find()) {
			p[i].pesudonym=m.group(0).trim().substring(1);
		}
	}
	public static void getBirthday(int i, String personinfo) {
		String pattern = "����.*?��";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(personinfo);
		if(m.find()) {
			String birthdayinfo=m.group(0).trim();
			p[i].ChineseBirthday = birthdayinfo.substring(2, birthdayinfo.length()-1);
		}
	}
	public static void getDeathdate(int i, String personinfo) {
		String pattern = "����.*?��";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(personinfo);
		if(m.find()) {
			String deathdateinfo=m.group(0).trim();
			p[i].Chinesedeathdate = deathdateinfo.substring(2, deathdateinfo.length()-1);
		}
	} 
	public static void getburied(int i, String personinfo) {
		if(personinfo.contains("��δ��")||personinfo.contains("��ʧ��")) {
			p[i].buried="ʧ��";
			return;
		}else {
			String pattern = "��[.*?]?��.*?��";
			Pattern r = Pattern.compile(pattern);
			Matcher m = r.matcher(personinfo);
			if(m.find()) {
				String buriedinfo=m.group(0).trim();
				p[i].buried = buriedinfo.substring(1);
			}
		}
	}
	public static void getRank(int i) {
		for(int j=12;j<generationRank.length;j++) {
			if(p[i].name.contains(generationRank[j])) {
				p[i].generition=j+1;
				break;
			}
		}
	}
	public static void getFatheridAndMotherid(int i) {
		for(int j=i-1;j>=0;j--) {
			if(p[j].name.equals(p[i].fathername)) {
				//System.out.println(p[j].toString());
				p[i].fatherid = p[j].id;
				int motherid = 0;
				if(p[j].partnerid!="") {
					if(p[j].partnerid.contains("/")) {
						motherid=Integer.parseInt(p[j].partnerid.split("/")[0]);
					}else {
						motherid=Integer.parseInt(p[j].partnerid);
					}
					p[i].motherid = motherid;
				}				
			}
		}
	}
	public static void getWifeINfo(int i, int num, String spouseinfo, String wiferank) {
		int idex = spouseinfo.indexOf('��');
		p[num] = new People();
		p[num].id = num+1;
		p[num].name = spouseinfo.substring(0, idex);
		p[num].gender="Ů";
		p[num].generition=p[i].generition;
		p[num].partnerid = String.valueOf(p[i].id);
		if(p[i].partnerid!="") {
			p[i].partnerid = p[i].partnerid +"/"+String.valueOf(p[num].id);
		}else {
			p[i].partnerid = String.valueOf(p[num].id);
		}
		getBirthday(num, spouseinfo);
    	getDeathdate(num, spouseinfo);
    	getburied(num, spouseinfo);
    	//System.out.println(p[num].toStringTest());
	}
	public static int getDaughterNum(String daughterInfo) {
		int num=0;
		String pattern = "Ů.";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(daughterInfo);
		if(m.find()) {
			switch (m.group(0)) {
				case "Ůһ":
					num=1;
					break;
				case "Ů��":
					num=2;
					break;
				case "Ů��":
					num=3;
					break;
				case "Ů��":
					num=4;
					break;
				case "Ů��":
					num=5;
					break;
				case "Ů��":
					num=6;
					break;
				case "Ů��":
					num=7;
					break;
				case "Ů��":
					num=8;
					break;
				case "Ů��":
					num=9;
					break;
				case "Ůʮ":
					num=10;
					break;
				default:
					num=0;
					break;
			}
		}		
		return num;
	}
	public static void addExcel(String path, List<People> list) {		
		//id+","+name+","+fatherid+","+motherid+","+partnerid+","+CourtesyName+","+pesudonym+","+gender+","+ChineseBirthday+","+Chinesedeathdate+","
				//+buried+","+familyrank+","+generition+","+family+","+description;
		try {
			WritableWorkbook wb = null;		
			// ������д���Excel������		
			File file = new File(path);
			if (!file.exists()) {
				file.createNewFile();
			}
			// ��fileNameΪ�ļ���������һ��Workbook
			wb = Workbook.createWorkbook(file);
			// ����������
			WritableSheet ws = wb.createSheet("Sheet0", 0);
			// Ҫ���뵽��Excel�����кţ�Ĭ�ϴ�0��ʼ
			Label labelId = new Label(0, 0, "���");
			Label labelName = new Label(1, 0, "����");
			Label labelfatherid = new Label(2,0,"���ױ��");
			Label labelmotherid = new Label(3,0,"ĸ�ױ��");
			Label labelpartnerid = new Label(4,0,"��ż���");
			Label labelcourtesyName = new Label(5,0,"��");
			Label labelpesudonym = new Label(6,0,"��");
			Label labelgender = new Label(7,0,"�Ա�");
			Label labelChineseBirthday = new Label(8,0,"ũ����������");
			Label labelChinesedeathdate = new Label(9,0,"ũ����������");
			Label labelburied = new Label(10,0,"����");
			Label labelfamilyrank = new Label(11,0,"��ͥ����");
			Label labelgenerition = new Label(12,0,"����");
			Label labelfamily = new Label(13,0,"��������");
			Label labeldescription = new Label(14,0,"��������");
			ws.addCell(labelId);
			ws.addCell(labelName);
			ws.addCell(labelfatherid);
			ws.addCell(labelmotherid);
			ws.addCell(labelpartnerid);
			ws.addCell(labelcourtesyName);
			ws.addCell(labelpesudonym);
			ws.addCell(labelgender);
			ws.addCell(labelChineseBirthday);
			ws.addCell(labelChinesedeathdate);
			ws.addCell(labelburied);
			ws.addCell(labelfamilyrank);
			ws.addCell(labelgenerition);
			ws.addCell(labelfamily);
			ws.addCell(labeldescription);
			for (int i = 0; i < list.size(); i++) {
				Label labelId_i = new Label(0, i + 1, list.get(i).id + "");
				Label labelName_i = new Label(1, i + 1, list.get(i).name);
				Label labelfatherid_i = new Label(2, i + 1, String.valueOf(list.get(i).fatherid));
				Label labelmotherid_i = new Label(3, i + 1, String.valueOf(list.get(i).motherid));
				Label labelpartnerid_i = new Label(4, i + 1, list.get(i).partnerid);
				Label labelcourtesyName_i = new Label(5, i + 1, list.get(i).CourtesyName);
				Label labelpesudonym_i = new Label(6, i + 1, list.get(i).pesudonym);
				Label labelgender_i = new Label(7, i + 1, list.get(i).gender);
				Label labelChineseBirthday_i = new Label(8, i + 1, list.get(i).ChineseBirthday);
				Label labelChinesedeathdate_i = new Label(9, i + 1, list.get(i).Chinesedeathdate);
				Label labelburied_i = new Label(10, i + 1, list.get(i).buried);
				Label labelfamilyrank_i = new Label(11, i + 1, String.valueOf(list.get(i).familyrank));
				Label labelgenerition_i = new Label(12, i + 1, String.valueOf(list.get(i).generition));
				Label labelfamily_i = new Label(13, i + 1, list.get(i).family);
				Label labeldescription_i = new Label(14, i + 1, list.get(i).description);
				ws.addCell(labelId_i);
				ws.addCell(labelName_i);
				ws.addCell(labelfatherid_i);
				ws.addCell(labelmotherid_i);
				ws.addCell(labelpartnerid_i);
				ws.addCell(labelcourtesyName_i);
				ws.addCell(labelpesudonym_i);
				ws.addCell(labelgender_i);
				ws.addCell(labelChineseBirthday_i);
				ws.addCell(labelChinesedeathdate_i);
				ws.addCell(labelburied_i);
				ws.addCell(labelfamilyrank_i);
				ws.addCell(labelgenerition_i);
				ws.addCell(labelfamily_i);
				ws.addCell(labeldescription_i);
			}
			// д���ĵ�
			wb.write();
			// �ر�Excel����������
			wb.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
