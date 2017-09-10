
package mailsender;
import java.util.*;
import java.io.*;
import javax.mail.*;
import javax.mail.internet.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class MailSender {

  private static final String FILE_PATH = "/home/rahul/Desktop/Untitled 1.xls";
  
     public void sendEmail(String host, String port,
			final String userName, final String password, String toAddress,
			String subject, String message) throws AddressException,
			MessagingException {

		// sets SMTP server properties
		Properties properties = new Properties();
		properties.put("mail.smtp.host", host);
		properties.put("mail.smtp.port", port);
		properties.put("mail.smtp.auth", "true");
		properties.put("mail.smtp.starttls.enable", "true");

		// creates a new session with an authenticator
		Authenticator auth = new Authenticator() {
			public PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(userName, password);
			}
		};

		Session session = Session.getInstance(properties, auth);

		// creates a new e-mail message
		Message msg = new MimeMessage(session);

		msg.setFrom(new InternetAddress(userName));
		InternetAddress[] toAddresses = { new InternetAddress(toAddress) };
		msg.setRecipients(Message.RecipientType.TO, toAddresses);
		msg.setSubject(subject);
		msg.setSentDate(new Date());
		// set plain text message
		msg.setText(message);

		// sends the e-mail
		Transport.send(msg);

	}
    
  public static void main(String[] args) 
    {
         FileInputStream file = null;
         String mail;
         String name;
                String host = "smtp.gmail.com";
		String port = "587";
		String mailFrom = "xyz@gmail.com";    //email_id
		String password = "pass";   //password
                String subject = "Subject";      
		String message = "Hello,";

        MailSender mailer = new MailSender();
         
        try {
            file = new FileInputStream(FILE_PATH);

            // Using XSSF for xlsx format, for xls use HSSF
            Workbook workbook = new HSSFWorkbook(file);

            int n = workbook.getNumberOfSheets();

            //looping over each workbook sheet
            for (int i = 0; i < n; i++) 
            {
                Sheet sheet = workbook.getSheetAt(i);
                
                Iterator<Row> rowIterator = sheet.iterator();
                
                rowIterator.next(); //skipping the first row
                
                while (rowIterator.hasNext()) {
                    
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    while (cellIterator.hasNext()) {

                        Cell cell = cellIterator.next();
                        
                        if (cell.getColumnIndex() == 1) {                          
                               mail= cell.getStringCellValue();
                              // System.out.print(mail);
                              
                              try {
			mailer.sendEmail(host, port, mailFrom, password, mail,
					subject, message);
			System.out.println("Email sent."+ mail);
		} catch (Exception ex) {
			System.out.println("Failed to sent email."+ mail);
			ex.printStackTrace();
		}
                             // System.out.print();
                        }
                        
                    }
                }
            }
            file.close();
        }
         catch (Exception e) {
            e.printStackTrace();
        }
    }
    
}
