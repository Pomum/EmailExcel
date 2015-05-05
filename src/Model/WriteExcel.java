package Model;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class WriteExcel 
{
	private String inputFile;
	private List<Email> deelnemersEmails;

	public WriteExcel(List<Email> deelnemersEmails) throws WriteException, IOException 
	{
		this.deelnemersEmails = deelnemersEmails;
		this.setOutputFile("c:/Excel/InschrijvingenDigitaal.xls");
		this.writeExcel();
		System.out.println("Please check the result file under c:/Excel/InschrijvingenDigitaal.xls");
	}
	
	public void setOutputFile(String inputFile) 
	{
		this.inputFile = inputFile;
	}

	public void writeExcel() throws IOException, WriteException 
	{
		File file = new File(inputFile);
		WorkbookSettings wbSettings = new WorkbookSettings();

		wbSettings.setLocale(new Locale("en", "EN"));

		WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
		workbook.createSheet("Deelnemers", 0);
		WritableSheet excelSheet = workbook.getSheet(0);

		writeLabels(excelSheet);
		writeContent(excelSheet);
		autoSize(excelSheet);
		
		workbook.write();
		workbook.close();
	}
	

	public void writeContent(WritableSheet excelSheet) throws RowsExceededException, WriteException
	{
		for(int e = 0; e < deelnemersEmails.size(); e++)
		{
			String s = deelnemersEmails.get(e).getMessage();
			String[] words = s.split("\\s+");
			int nummer = e + 1;
			excelSheet.addCell(new Label(0, e+1, "" + nummer));
			excelSheet.addCell(new Label(1, e+1, "Digitaal"));
			
			for(int i = 0; i < words.length; i++)
			{
				if(words[i].equals("Voornaam:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Tussenvoegsel:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(3, e+1, str));
				}
				if(words[i].equals("Tussenvoegsel:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Achternaam:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(4, e+1, str));
				}
				if(words[i].equals("Achternaam:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Straat:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(5, e+1, str));
				}
				if(words[i].equals("Straat:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Postcode:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(6, e+1, str));
				}
				if(words[i].equals("Postcode:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Plaats:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(7, e+1, str));
				}
				if(words[i].equals("Plaats:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Tel:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(8, e+1, str));
				}
				if(words[i].equals("Tel:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("E-mail:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(9, e+1, str));
				}
				if(words[i].equals("E-mail:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Geboortedatum:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(10, e+1, str));
				}
				if(words[i].equals("Geboortedatum:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Basisschool:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(11, e+1, str));
				}
				if(words[i].equals("Basisschool:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Huisarts:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(12, e+1, str));
				}
				if(words[i].equals("Huisarts:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Allergieën:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(13, e+1, str));
				}
				if(words[i].equals("Allergieën:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Bijzonderheden:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(14, e+1, str));
				}
				if(words[i].equals("Bijzonderheden:"))
				{
					i++;
					String str = "";
					while(!words[i].equals("Group:"))
					{
						str += words[i] + " ";
						i++;
					}
					excelSheet.addCell(new Label(15, e+1, str));
				}
				if(words[i].equals("Group:"))
				{
					i++;
					if(words[i].equals("no"))
					{
						excelSheet.addCell(new Label(16, e+1, "Nee"));
					} 
					else if(words[i].equals("yes"))
					{
						excelSheet.addCell(new Label(16, e+1, "Ja"));
						i++;
						String str = "";
						if(words[i].equals("Eerste:"))
						{
							i++;
							
							while(!words[i].equals("Tweede:"))
							{
								str += words[i] + " ";
								i++;
							}
							
						}
						if(words[i].equals("Tweede:"))
						{
							i++;
							while(!words[i].equals("Derde:"))
							{
								str += words[i] + " ";
								i++;
							}
							
						}
						if(words[i].equals("Derde:"))
						{
							i++;
							while(!words[i].equals("Vierde:"))
							{
								str += words[i] + " ";
								i++;
							}
							
						}
						if(words[i].equals("Vierde:"))
						{
							i++;
							while(!words[i].equals("Vijfde:"))
							{
								str += words[i] + " ";
								i++;
							}
							
						}
						if(words[i].equals("Vijfde:"))
						{
							i++;
							for(int j = i; j < words.length; j++)
							{
								str += words[j] + " ";
							}
						}
						excelSheet.addCell(new Label(17, e+1, str));
					}
				}
			}
		}
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
	
	private void autoSize(WritableSheet excelSheet) 
	{
		CellView cf = new CellView();
	    cf.setAutosize(true);
	    for(int i = 0; i < 18; i++)
	    {
	    	if(i != 15)
	    	{
	    		excelSheet.setColumnView(i, cf);
	    	}
	    }
	}
}