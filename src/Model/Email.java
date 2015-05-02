package Model;

public class Email 
{
	private String subject;
	private String sender;
	private String message;
	
	public Email(String subject, String sender, String message)
	{
		this.subject = subject;
		this.sender = sender;
		this.message = message;
	}
	
	public String getSubject()
	{
		return subject;
	}
	
	public String getSender()
	{
		return sender;
	}
	
	public String getMessage()
	{
		return message;
	}
}
