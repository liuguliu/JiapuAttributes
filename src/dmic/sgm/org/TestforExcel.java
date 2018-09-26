package dmic.sgm.org;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;


public class TestforExcel {
	public static void main(String[] args) {
		String path = "data/test.xls";
		People people = new People();
		people.id = 1;
		people.name = "me";
		people.gender = "男";
		
		List<People> list = new ArrayList<People>();
		list.add(people);
		addExcel(path, list);
	}
	public static void addExcel(String path, List<People> list) {		
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
			WritableSheet ws = wb.createSheet("Test Shee 1", 0);
			// 要插入到的Excel表格的行号，默认从0开始
			Label labelId = new Label(0, 0, "编号");// 表示第
			Label labelName = new Label(1, 0, "姓名");
			Label labelPwd = new Label(2, 0, "密码");
			ws.addCell(labelId);
			ws.addCell(labelName);
			ws.addCell(labelPwd);
			for (int i = 0; i < list.size(); i++) {
				Label labelId_i = new Label(0, i + 1, list.get(i).id + "");
				Label labelName_i = new Label(1, i + 1, list.get(i).name);
				Label labelSex_i = new Label(2, i + 1, list.get(i).gender);
				ws.addCell(labelId_i);
				ws.addCell(labelName_i);
				ws.addCell(labelSex_i);
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
