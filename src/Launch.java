import java.io.IOException;

import jxl.write.WriteException;
import Controller.EmailController;

/** 
 * Documentation: https://javamail.java.net/nonav/docs/api/  
 * */

public class Launch 
{
	public static void main(String[] args) throws WriteException, IOException 
	{
		new EmailController();
	}
}
