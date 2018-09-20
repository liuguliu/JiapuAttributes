package dmic.sgm.org;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ooxml.extractor.POIXMLTextExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class TestGit {
	public static void main(String[] args) {
		String text="";
		String filePath = "data";
		String fileName = "12.doc";
        String realPath=filePath+"/"+fileName;//ƴ��Ϊ�����ֵ�·��
        File file = new File(realPath);
        if(file.exists()) {
        	System.out.println("YES");
        }
        System.out.println(realPath);
        try {
            if(fileName.endsWith(".doc")){   //docΪ��׺��
                //FileInputStream in= new FileInputStream(realPath);
                //WordExtractor extractor = new WordExtractor(in);
                InputStream is = new FileInputStream(realPath);  
                //WordExtractor extractor = new WordExtractor(is);  
                HWPFDocument extractor = new HWPFDocument(is); 
                //���word�ĵ����е��ı�  
                System.out.println(extractor.getText()); 
                //text = extractor.getText();
            }
            if(fileName.endsWith(".docx")){  //docxΪ��׺��
            	OPCPackage oPCPackage = POIXMLDocument.openPackage(filePath);
                XWPFDocument xwpf = new XWPFDocument(oPCPackage);
                POIXMLTextExtractor ex = new XWPFWordExtractor(xwpf);
                text = ex.getText();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        if(!"".equals(text)){
        	System.out.print("success!");
        }
	}
}
