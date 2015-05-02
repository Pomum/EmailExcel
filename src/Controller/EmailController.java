package Controller;

import java.util.ArrayList;
import java.util.List;

import Model.Email;
import Model.ReadMails;

public class EmailController 
{
	private List<Email> deelnemersMails = new ArrayList<Email>();
	private List<Email> vrijwilligersMails = new ArrayList<Email>();
	
	public EmailController()
	{
		new ReadMails(this);
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
