package Model;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelTest 
{
	private String inputFile;

	public static void main(String[] args) throws WriteException, IOException 
	{
		ExcelTest excel = new ExcelTest();
		excel.setOutputFile("c:/Excel/testExcel.xls");
		excel.write();
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

		writeLabels(excelSheet);
		
		workbook.write();
		workbook.close();
	}
	
	public void writeLabels(WritableSheet excelSheet) throws RowsExceededException, WriteException
	{
		excelSheet.addCell(new Label(0,0, "Inschrijfnummer:"));
		excelSheet.addCell(new Label(1,0, "d/f"));
		excelSheet.addCell(new Label(2,0, "Betaald:"));
		excelSheet.addCell(new Label(3,0, "Voornaam:"));
		excelSheet.addCell(new Label(4,0, "Tussenvoegsel:"));
		excelSheet.addCell(new Label(5,0, "Achternaam:"));
		excelSheet.addCell(new Label(6,0, "Straat:"));
		excelSheet.addCell(new Label(7,0, "P.C.:"));
		excelSheet.addCell(new Label(8,0, "Plaats:"));
		excelSheet.addCell(new Label(9,0, "Telnr:"));
		excelSheet.addCell(new Label(10,0, "E-mail:"));
		excelSheet.addCell(new Label(11,0, "Geb datum:"));
		excelSheet.addCell(new Label(12,0, "School:"));
		excelSheet.addCell(new Label(13,0, "Huisarts:"));
		excelSheet.addCell(new Label(14,0, "Allergien:"));
		excelSheet.addCell(new Label(15,0, "Bijzonderheden:"));
		excelSheet.addCell(new Label(16,0, "Groep:"));
		excelSheet.addCell(new Label(17,0, "Wie:"));
		excelSheet.addCell(new Label(18,0, "Groepsnr:"));
		
	}
}