package parse;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class GetNo {

	public static void main(String[] args) throws RowsExceededException, WriteException, IOException {
		ArrayList<String> als = new ArrayList<String>();
		ArrayList<Entry> alr = new ArrayList<Entry>();
		
		File f = new File("D:\\a.xls");
		try {
			Workbook book = Workbook.getWorkbook(f);//
			Sheet sheet = book.getSheet(0); // 获得第一个工作表对象
			for (int i = 0; i < sheet.getRows(); i++) {
				Entry e = new Entry();
				Cell cell = sheet.getCell(0, i); // 获得单元格
				e.name = cell.getContents();
				cell = sheet.getCell(1, i); // 获得单元格
				e.no = cell.getContents();
				alr.add(e);
				// System.out.print(cell.getContents() + " ");

				// System.out.print("\n");
			}
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println(alr.size());
		  f = new File("D:\\c.xls");
		
		try {
			Workbook book = Workbook.getWorkbook(f);//
			Sheet sheet = book.getSheet(0); // 获得第一个工作表对象
			for (int i = 0; i < sheet.getRows(); i++) {
				Cell cell = sheet.getCell(1, i); // 获得单元格
				als.add(cell.getContents());
				 //System.out.print(cell.getContents() + " ");

				// System.out.print("\n");
			}
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		f = new File("D:\\b.xls");
		try {
			Workbook book = Workbook.getWorkbook(f);//
			Sheet sheet = book.getSheet(0); // 获得第一个工作表对象
			for (int i = 0; i < sheet.getRows(); i++) {

				Cell cell = sheet.getCell(0, i); // 获得单元格
				als.add(cell.getContents());
				// System.out.print(cell.getContents() + " ");

				// System.out.print("\n");
			}
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

			WritableWorkbook book = null;
			WritableSheet sheet;
			Workbook wb = null; // 创建一个workbook对象/
			book=Workbook.createWorkbook(new File("E://BookInfo___" + Math.random()
					+ ".xls"));
			sheet = book.createSheet("A",1);
			
		int count = 0;
		for (int i = 0; i < als.size(); i++) {
	
		
			for (int j = 0; j < alr.size(); j++) {
				if (alr.get(j).no.equals(als.get(i))) {
System.out.println(alr.get(j).no+" "+als.get(i));
					
					Label label = new Label(0,count , alr.get(j).name);
					sheet.addCell(label);
					 label = new Label(1,count , alr.get(j).no);
					sheet.addCell(label);
					count++;
					// System.out.println(es.name+"  "+same.getSimilarityRatio(es.name,als.get(j)));
				}
			}

		}
	
		book.write();
		book.close();

	}

}
