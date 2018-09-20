package dmic.sgm.org;

public class People {
	int id;	
	String name;
	int fatherid;
	String partnerid;
	String CourtesyName;
	String pesudonym;
	String ChineseBirthday;
	String Chinesedeathdate;
	String buried; 
	int rank;
	int generition;
	String description;
	String fathername="";
	String daughterinfo="";
	String soninfo="";
	String spouseinfo="";
	String personInfo="";
	
	@Override
	public String toString() {
		// TODO Auto-generated method stub
		return id+","+name+","+fatherid+","+partnerid+","+CourtesyName+","+pesudonym+","+ChineseBirthday+","+Chinesedeathdate+","
				+buried+","+rank+","+generition+","+description;
	}
	
	public String toStringTest() {
		// TODO Auto-generated method stub
		return id+","+name+","+partnerid+","+CourtesyName+","+pesudonym+","+ChineseBirthday+","+Chinesedeathdate+","
				+buried+","+rank+","+generition+","+daughterinfo+","+soninfo+","+description;
	}
}
