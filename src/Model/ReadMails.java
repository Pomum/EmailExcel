package Model;
import java.io.IOException;
import java.util.Properties;

import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.NoSuchProviderException;
import javax.mail.Session;

import Controller.EmailController;

import com.sun.mail.pop3.POP3Store;


public class ReadMails 
{
	private EmailController ec;
	
	public ReadMails(EmailController ec)
	{
		this.ec = ec;
		
		String mailPop3Host = "";
		String mailStoreType = "pop3";
		String mailUser = "";
		String mailPassword = "";

		receiveEmail(mailPop3Host, mailStoreType, mailUser, mailPassword);
	}

	public void receiveEmail(String pop3Host, String storeType, String user, String password) 
	{
		try 
		{
			Properties properties = new Properties();
			properties.put("mail.pop3.host", pop3Host);
			Session emailSession = Session.getDefaultInstance(properties);

			POP3Store emailStore = (POP3Store) emailSession.getStore(storeType);
			emailStore.connect(user, password);

			Folder emailFolder = emailStore.getFolder("INBOX");
			emailFolder.open(Folder.READ_ONLY);

			Message[] messages = emailFolder.getMessages();
			for (Message message : messages) 
			{
				if(message.getSubject().equals("Inschrijving Deelnemers"))
				{
					Email email = new Email(message.getSubject(), "" + message.getFrom()[0], message.getContent().toString());
					ec.addDeelnemersMails(email);
				}
				else if(message.getSubject().equals("Inschrijving Vrijwilligers"))
				{
					Email email = new Email(message.getSubject(), "" + message.getFrom()[0], message.getContent().toString());
					ec.addVrijwilligersMails(email);
				}
				else 
				{
					System.out.println(message.getSubject());
				}
			}

			emailFolder.close(false);
			emailStore.close();
		} 
		catch (NoSuchProviderException e) 
		{
			e.printStackTrace();
		} catch (MessagingException e) 
		{
			e.printStackTrace();
		} catch (IOException e) 
		{
			e.printStackTrace();
		}
	}
}
