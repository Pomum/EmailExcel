package Model;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelTest 
{
	private WritableCellFormat arialBoldUnderline;
	private WritableCellFormat arial;
	private String inputFile;

	public static void main(String[] args) throws WriteException, IOException 
	{
		ExcelTest test = new ExcelTest();
		test.setOutputFile("c:/Excel/testExcel.xls");
		test.write();
		System.out.println("Please check the result file under c:/Excel/testExcel.xls");
	}
	
	public void setOutputFile(String inputFile) 
	{
		this.inputFile = inputFile;
	}

	public void write() throws IOException, WriteException 
	{
		File file = new File(inputFile);
		WorkbookSettings wbSettings = new WorkbookSettings();

		wbSettings.setLocale(new Locale("en", "EN"));

		WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
		workbook.createSheet("Deelnemers", 0);
		WritableSheet excelSheet = workbook.getSheet(0);
		workbook.createSheet("Vrijwilligers", 1);
		WritableSheet excelSheet2 = workbook.getSheet(1);

		workbook.write();
		workbook.close();
	}
}