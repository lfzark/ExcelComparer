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


/**
* @author Artk
* @version Create Time：2014-6-18 下午7:33:10
* @Use 
*/
public class Compare {

	public static void main(String[] args) throws IOException,
			RowsExceededException, WriteException {
		Same same = new Same();
		ArrayList<Entry> ale = new ArrayList<Entry>();
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
				ale.add(e);
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
		System.out.println(ale.size());
		System.out.println(als.size());

		WritableWorkbook book = null;
		WritableSheet sheet;
		//Workbook wb = null;  
		book = Workbook.createWorkbook(new File("E:/BookInfo" + Math.random()
				+ ".xls"));
		sheet = book.createSheet("A", 1);

		for (int i = 0; i < 530; i++) {
			Entry es = ale.get(i);

			for (int j = 0; j < 641; j++) {
				if (same.getSimilarityRatio(es.name, als.get(j)) > 0.5
						|| (es.name.contains(als.get(j)) && als.get(j).length() > 5)) {

					alr.add(es);
					System.out.println(es.name + "   " + als.get(j) + "  "
							+ es.no + " "
							+ same.getSimilarityRatio(es.name, als.get(j)));
					Label label = new Label(0, alr.size(), es.name);
					sheet.addCell(label);
					label = new Label(1, alr.size(), als.get(j));
					sheet.addCell(label);
					label = new Label(2, alr.size(), es.no);
					sheet.addCell(label);
					label = new Label(3, alr.size(), String.valueOf(same
							.getSimilarityRatio(es.name, als.get(j))));
					sheet.addCell(label);

					// System.out.println(es.name+"  "+same.getSimilarityRatio(es.name,als.get(j)));
				}
			}

		}

		book.write();
		book.close();
		System.out.println(alr.size());

	}
}
