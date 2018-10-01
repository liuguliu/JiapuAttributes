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
	// 定义静态变量
	static People p[] = new People[1000];
	static String cNum[] = {"一","二","三","四","五","六","七","八","九","十"};
	static String rankNum[] = {"","长","次","三","四","五","六","七","八","九","十"};
	static String generationRank[] = {"一","二","三","四","五","六","七","八","九","十","十一","十二","延","祚","昌","克","相","盛","时"};
	public static void main(String[] args) {
		
		//String[] commonstring = {"生未详","生卒未详","公卒葬失考","妣卒葬失考","妣卒葬未详","公妣卒葬失考","公与王妣卒未详","生卒葬失考"};
		
		String text="";
		String filePath = "data";
		String fileName = "12.doc";
        String realPath=filePath+"/"+fileName;	//拼接为含名字的路径
        try {
        	if(fileName.endsWith(".doc")) {
        		InputStream is = new FileInputStream(realPath);    
                HWPFDocument extractor = new HWPFDocument(is); 
                //输出word文档所有的文本  
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
        for(int i=0; i<lines.length; i++) {	//1 循环读取每一行的人物
        	//System.out.println(lines[i]);
        	p[i] = new People();	//新建人物
        	p[i].id = i+1; 
        	p[i].description = lines[i];
//        	//2 提取配偶信息
//        	int indexofpei;
//        	if(lines[i].contains("配")) {
//        		indexofpei=lines[i].indexOf("配");
//        	}else if(lines[i].contains("妣")){
//        		indexofpei=lines[i].indexOf("配");
//        	}else {
//        		indexofpei = -1;
//        	}
//        	int indexofend;
//        	if(lines[i].contains(" 子")) {
//        		indexofend = lines[i].indexOf(" 子");
//        	}
//        	else if(lines[i].contains(" 女")) {
//        		indexofend = lines[i].indexOf(" 女")+1;
//        	}
//        	else if(lines[i].contains(" 公")) {
//        		indexofend = lines[i].indexOf(" 公")+1;
//        	}
//        	else {
//        		indexofend = lines[i].length();
//        	}
//        	//System.out.println(indexofpei+" "+indexofend);
//        	if(indexofpei!=-1) {
//        		p[i].spouseinfo+=lines[i].substring(indexofpei, indexofend);        		
//        	}
//        	if(lines[i].contains("。妣")) {
//        		int indexbi = lines[i].indexOf("妣");
//        		p[i].spouseinfo+=lines[i].substring(indexbi+1);
//        	}
//        	if(p[i].spouseinfo!="") {
//        		String[] couple = p[i].spouseinfo.split("。");
//            	p[i].spouseinfo="";
//            	for(int j=0;j<couple.length;j++) {	//清除公妣共有信息
//            		if(couple[j].contains("公")&&couple[j].contains("妣")) {
//            			break;
//            		}
//            		else {
//            			p[i].spouseinfo+=couple[j]+"。";
//            		}
//            	} 
//        	}     	      	
        	
        	String[] lineAttrs = lines[i].split("。");
        	for(int j=0;j<lineAttrs.length;j++) {	//3 将人物属性与配偶信息及子女信息分开
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
        	//System.out.println(p[i].personInfo);
        	//System.out.println(p[i].spouseinfo);
        	//System.out.println(p[i].soninfo);
        	//System.out.println(p[i].daughterinfo);
        	//System.out.println("完");
        	//4 提取人物属性
        	getName(i, p[i].personInfo);
        	getFatherNameAndRank(i, p[i].personInfo);
        	getCourtesyName(i, p[i].personInfo);
        	getpesudonym(i, p[i].personInfo);
        	getBirthday(i, p[i].personInfo);
        	getDeathdate(i, p[i].personInfo);
        	getburied(i, p[i].personInfo);
        	getRank(i);
        	//System.out.println(p[i].toStringTest());
        	//5 处理配偶信息
        	if(p[i].spouseinfo!="") {
        		if(p[i].spouseinfo.contains("继配")) {
        			String wife2 = p[i].spouseinfo.substring(p[i].spouseinfo.indexOf("继配")+2);
            		String wife1 = p[i].spouseinfo.substring(1, p[i].spouseinfo.indexOf("继配"));
            		getWifeINfo(i, num, wife1, "");
            		num++;           		
            		if(p[i].spouseinfo.contains("三配")) {
            			 String wife3 = wife2.substring(wife2.indexOf("三配")+2);
            			 wife2 = wife2.substring(0,wife2.indexOf("三配"));
            			 getWifeINfo(i, num, wife2, "继配");
                 		 num++;
                 		 getWifeINfo(i, num, wife3, "三配");
                		 num++;
            		}else {
            			 getWifeINfo(i, num, wife2, "继配");
                		 num++;
            		}
            	}else {
            		getWifeINfo(i, num, p[i].spouseinfo.substring(1), "");
            		num++;
            	}
        	}
        	getFatheridAndMotherid(i);
        }	//for
        for(int i=0; i<lines.length; i++) {	//1 循环读取每一行的人物
        	//6 处理儿子信息
        	if(p[i].soninfo!="") {
        		//System.out.println(p[i].soninfo);
        		String[] sons = p[i].soninfo.split(" ");
        		for(int j =1; j<sons.length;j++) {
        			String nameson="";
        			String description = "";
        			//提取姓名和描述
        			if(sons[j].contains("(")) {
        				nameson=sons[j].substring(0, sons[j].indexOf("("));
        				description = sons[j].substring(sons[j].indexOf("(")+1,sons[j].length()-1);
        				//System.out.print(nameson+" ");
        			}else{
        				nameson = sons[j];
        				//System.out.print(nameson+" ");
        			}
        			//判断是否存在
        			boolean flag=false;
        			for(int k=0;k<lines.length;k++) {
        				if((p[k].name.equals(nameson)||p[k].name.equals(generationRank[p[i].generition]+nameson))
        						&&p[k].fathername.equals(p[i].name)) {
        					flag=true;
        					break;
        				}
        			}
        			if(!flag) {	//若不存在
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
        	if(p[i].gender.equals("男")) {
        		p[i].name = "吴" + p[i].name;
        	}
        }
        for(int i=0;i<num;i++) {    	
        	//7 处理女儿信息
        	if(p[i].daughterinfo!="") {
        		System.out.println(p[i].daughterinfo);
        		int daughternum = getDaughterNum(p[i].daughterinfo);
        		if(!p[i].daughterinfo.contains(" ")) {   //俱夭     			        			
    				for(int j=0;j<daughternum;j++) {
    					p[num] = new People();
    					p[num].id = num+1;
    					p[num].name = "女"+cNum[j];
    					p[num].gender="女";
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
        					p[num].description = "夭";
        				}
        				//System.out.println(p[num].toString());
        				num++;
    				}
        		}else {
        			String[] daughters = p[i].daughterinfo.split(" ");
        			for(int j=1;j<daughternum;j++) {
        				if(daughters[j].contains(rankNum[j]+"适")||daughters[j].charAt(0)=='适') {
        					p[num]=new People();
        					p[num].id=num+1;
        					p[num].name="女"+cNum[j-1];
        					p[num].gender="女";
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
            				if(daughters[j].contains(rankNum[j]+"适")) {
            					p[num].description = daughters[j].substring(1);
            				}else {
            					p[num].description = daughters[j];
            				}
            				p[num].partnerid = String.valueOf(num+2);
            				//System.out.println(p[num].toString());
            				num++;
            				//嫁
            				p[num]=new People();
        					p[num].id=num+1;
        					if(daughters[j].contains(rankNum[j]+"适")) {
        						p[num].name=daughters[j].substring(2);
            				}else {
            					p[num].name=daughters[j].substring(1);
            				}
        					p[num].partnerid = String.valueOf(num);
        					num++;
        				}else  {
        					p[num]=new People();
        					p[num].id=num+1;
        					p[num].name="女"+cNum[j-1];
        					p[num].gender="女";
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
            				if(daughters[j].contains("(殀)")||daughters[j].contains("(夭)")) {
            					p[num].description = "夭";
            				}
            				//System.out.println(p[num].toString());
            				num++;
        				}
        			}
        		}
        	}//7 处理女儿信息
        }
        String path = "data/第15卷.xls";
        List<People> list = new ArrayList<People>();
//        for(int i=0;i<num;i++) {
//        	list.add(p[i]);
//        	System.out.println(p[i].toString());
//        }
        //addExcel(path,list);
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
	public static void getFatherNameAndRank(int i, String personinfo) {
		int indexStart = personinfo.indexOf(' ');
		int indexEnd = personinfo.indexOf("子 ");
		
		String fatherAndRank = personinfo.substring(indexStart, indexEnd+1);
		//System.out.println(indexStart+" "+indexEnd+" "+fatherAndRank);
		int indexofRank = getFamilyRank(i, fatherAndRank);
		//System.out.println(indexofRank);
		p[i].fathername = fatherAndRank.substring(1,indexofRank);

	}
	public static int getFamilyRank(int i, String fartherAndRank) {
		int indexofRank = -1;
		
		if(fartherAndRank.contains("之子")||fartherAndRank.contains("长子")) {
			p[i].familyrank=1;
			if(fartherAndRank.indexOf("之子")!=-1) {
				indexofRank = fartherAndRank.indexOf("之子");
			}
			else {
				indexofRank = fartherAndRank.indexOf("长子");
			}
		}
		else if(fartherAndRank.contains("次子")) {
			p[i].familyrank=2;
			indexofRank = fartherAndRank.indexOf("次子");
		}
		else if(fartherAndRank.contains("三子")) {
			p[i].familyrank=3;
			indexofRank = fartherAndRank.indexOf("三子");
		}
		else if(fartherAndRank.contains("四子")) {
			p[i].familyrank=4;
			indexofRank = fartherAndRank.indexOf("四子");
		}
		else if(fartherAndRank.contains("五子")) {
			p[i].familyrank=5;
			indexofRank = fartherAndRank.indexOf("五子");
		}
		else if(fartherAndRank.contains("六子")) {
			p[i].familyrank=6;
			indexofRank = fartherAndRank.indexOf("六子");
		}
		else if(fartherAndRank.contains("七子")) {
			p[i].familyrank=7;
			indexofRank = fartherAndRank.indexOf("七子");
		}
		else if(fartherAndRank.contains("八子")) {
			p[i].familyrank=8;
			indexofRank = fartherAndRank.indexOf("八子");
		}
		else if(fartherAndRank.contains("九子")) {
			p[i].familyrank=9;
			indexofRank = fartherAndRank.indexOf("九子");
		}
		else if(fartherAndRank.contains("十子")) {
			p[i].familyrank=10;
			indexofRank = fartherAndRank.indexOf("十子");
		}
		return indexofRank;
	}
	public static void getCourtesyName(int i, String personinfo) {
		String pattern = "字.*? ";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(personinfo);
		if(m.find()) {
			String courtesyinfo =m.group(0).trim();
			if(courtesyinfo.contains("又字")) {
				int indexofYou = courtesyinfo.indexOf("又字");
				p[i].CourtesyName = courtesyinfo.substring(1, indexofYou);
				p[i].CourtesyName = p[i].CourtesyName + "/"+courtesyinfo.substring(indexofYou+2);
			}
			else {
				p[i].CourtesyName = courtesyinfo.substring(1);
			}
		}
	}
	public static void getpesudonym(int i, String personinfo) {
		String pattern = "号.*? ";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(personinfo);
		if(m.find()) {
			p[i].pesudonym=m.group(0).trim().substring(1);
		}
	}
	public static void getBirthday(int i, String personinfo) {
		String pattern = "生于.*?。";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(personinfo);
		if(m.find()) {
			String birthdayinfo=m.group(0).trim();
			p[i].ChineseBirthday = birthdayinfo.substring(2, birthdayinfo.length()-1);
		}
	}
	public static void getDeathdate(int i, String personinfo) {
		String pattern = "卒于.*?。";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(personinfo);
		if(m.find()) {
			String deathdateinfo=m.group(0).trim();
			p[i].Chinesedeathdate = deathdateinfo.substring(2, deathdateinfo.length()-1);
		}
	} 
	public static void getburied(int i, String personinfo) {
		if(personinfo.contains("葬未详")||personinfo.contains("葬失考")) {
			p[i].buried="失考";
			return;
		}else {
			String pattern = "。[.*?]?葬.*?。";
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
		int idex = spouseinfo.indexOf('。');
		p[num] = new People();
		p[num].id = num+1;
		p[num].name = spouseinfo.substring(0, idex);
		p[num].gender="女";
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
		String pattern = "女.";
		Pattern r = Pattern.compile(pattern);
		Matcher m = r.matcher(daughterInfo);
		if(m.find()) {
			switch (m.group(0)) {
				case "女一":
					num=1;
					break;
				case "女二":
					num=2;
					break;
				case "女三":
					num=3;
					break;
				case "女四":
					num=4;
					break;
				case "女五":
					num=5;
					break;
				case "女六":
					num=6;
					break;
				case "女七":
					num=7;
					break;
				case "女八":
					num=8;
					break;
				case "女九":
					num=9;
					break;
				case "女十":
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
			// 创建可写入的Excel工作簿		
			File file = new File(path);
			if (!file.exists()) {
				file.createNewFile();
			}
			// 以fileName为文件名来创建一个Workbook
			wb = Workbook.createWorkbook(file);
			// 创建工作表
			WritableSheet ws = wb.createSheet("Sheet0", 0);
			// 要插入到的Excel表格的行号，默认从0开始
			Label labelId = new Label(0, 0, "编号");
			Label labelName = new Label(1, 0, "姓名");
			Label labelfatherid = new Label(2,0,"父亲编号");
			Label labelmotherid = new Label(3,0,"母亲编号");
			Label labelpartnerid = new Label(4,0,"配偶编号");
			Label labelcourtesyName = new Label(5,0,"字");
			Label labelpesudonym = new Label(6,0,"号");
			Label labelgender = new Label(7,0,"性别");
			Label labelChineseBirthday = new Label(8,0,"农历出生日期");
			Label labelChinesedeathdate = new Label(9,0,"农历过世日期");
			Label labelburied = new Label(10,0,"葬于");
			Label labelfamilyrank = new Label(11,0,"家庭排行");
			Label labelgenerition = new Label(12,0,"辈份");
			Label labelfamily = new Label(13,0,"所属姓氏");
			Label labeldescription = new Label(14,0,"其他描述");
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
			// 写进文档
			wb.write();
			// 关闭Excel工作簿对象
			wb.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
