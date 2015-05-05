package Controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import jxl.write.WriteException;
import Model.Email;
import Model.ReadMails;
import Model.WriteExcel;

public class EmailController 
{
	private List<Email> deelnemersMails = new ArrayList<Email>();
	private List<Email> vrijwilligersMails = new ArrayList<Email>();
	
	public EmailController() throws WriteException, IOException
	{
		new ReadMails(this);
		new WriteExcel(deelnemersMails);
	}
	
	public void addDeelnemersMails(Email mail)
	{
		deelnemersMails.add(mail);
	}
	
	public void addVrijwilligersMails(Email mail)
	{
		vrijwilligersMails.add(mail);
	}
}
